"""Extract files from OLE object fields inside Microsoft Access databases"""

import struct
import re
import sys
from pprint import pprint

class BadDataError(Exception):
    pass

def bmps(oleobject):
    """Return iterator over all BMPs inside OLE object field"""

    for object_type, data in objects(oleobject):
        if object_type == 'METAFILEPICT':
            for image in metafile_bmps(data):
                yield image
        elif object_type == 'PBrush':
            # PBrush data is already BMP
            yield data

def objects(oleobject, verbose=False):
    """Parse OLE object field and return iterator over 'objects' embedded

    Each iteration returns (object_type, object_content) tuple
    """

    # The best information about OLE objects in Microsoft Access fields:
    # * http://jvdveen.blogspot.com/2009/01/ole-and-accessing-files-embedded-in.html
    # * http://jvdveen.blogspot.com/2009/02/ole-and-accessing-files-embedded-in.html

    s = oleobject

    # oleobject field header
    length, header = unwrap(s, """h signature == 0x1c15 !
                                  h header_size
                                  i object_type
                                  h friendly_len
                                  h class_len
                                  h friendly_off
                                  h class_off""")
    if verbose: pprint(header)

    def unpack_name(what):
        offset, length = header['%s_off' % what], header['%s_len' % what]
        if offset + length > header['header_size']:
            raise BadDataError('Bad header')
        return s[offset:offset+length].rstrip('\x00')

    names = dict((n, unpack_name(n)) for n in ('friendly', 'class'))
    if verbose: print names

    s = s[header['header_size']:]

    while 1:
        if len(s) <= 4:
            length, footer = unwrap(s, """1s unknown
                                          3s footer == '\xad\x05\xfe' !
                                       """)
            if verbose: print 'footer %r' % footer
            break

        # Seems to be mostly in OLE 1.0 format, documented in "[MS-OLEDS]:
        # Object Linking and Embedding (OLE) Data Structures", section 2.2
        # http://msdn.microsoft.com/en-us/library/dd942265.aspx
        length, ole_header = unwrap(s, """I ole_version == 0x0501 ?
                                          I ole_format
                                          i object_type_len""")
        if ole_header['ole_format'] == 0: # empty object
            s = s[8:] # skip it (the rest is usually footer)
            continue

        if verbose: pprint(ole_header)
        s = s[length:]

        length, ole_header_cont = unwrap(s, """%(object_type_len)ss object_type
                                               8s unknown""" % ole_header)
        # Observations about ole_header_cont['unknown']:
        #  object_type=METAFILEPICT: [ii] bmp_width*~26.46, -bmp_height*~26.46
        #                                 (confirmed by [MS-OLEDS], 2.2.2)
        #  object_type=PBrush: all zeros
        if verbose: pprint(ole_header_cont)
        s = s[length:]

        length, block_header = unwrap(s, """i data_block_len""")

        if verbose: pprint(block_header)

        data = s[length:length+block_header['data_block_len']]
        #if verbose: print 'data', sformat(data)
        yield ole_header_cont['object_type'].rstrip('\x00'), data

        s = s[length+block_header['data_block_len']:]

META_EOF = 0x0000
META_DIBSTRETCHBLT = 0x0b41
BITMAPINFOHEADER = 40
BI_BITCOUNT_5 = 0x0018

def metafile_bmps(metafilepict, verbose=False):
    """Parse OLE field METAFILEPICT objects, return iterator over BMPs"""

    # METAFILEPICT object is a Windows Metafile, documented in
    # "[MS-WMF]: Windows Metafile Format (wmf) Specification",
    # http://msdn.microsoft.com/en-us/library/cc215212.aspx

    s = metafilepict

    # Eight reserved bytes ([MS-OLEDS], 2.2.2.1), then WMF header ([MS-WMF],
    # 2.3.2.2). In practice first two bytes were always 0x0008 for me, the
    # rest six are most certainly garbage.
    length, header = unwrap(s, """8s reserved
                                  H type == 0x0001 ?
                                  H header_size
                                  H version == 0x0300 ?
                                  I metafile_size
                                  H num_of_objects == 0 ?
                                  I max_record_len
                                  H unused_should_be_0""", 'metafile')
    if verbose: pprint(header)
    s = s[length:]

    while s:
        # metafile record ([MS-WMF], 2.3)
        length, record_header = unwrap(s, """I record_size
                                             H function""")
        if verbose: pprint(record_header)

        if record_header['function'] == META_DIBSTRETCHBLT:
            # META_DIBSTRETCHBLT record parameters ([MS-WMF], 2.3.1.3)
            L, blt_header = unwrap(s[length:], """I raster_operation
                                                  h src_height
                                                  h src_width
                                                  h y_src
                                                  h x_src
                                                  h dest_height
                                                  h dest_width
                                                  h y_dest
                                                  h x_dest""")
            if verbose: pprint(blt_header)

            # ensure record has bitmap (directly from [MS-WMF], 2.3.1.3)
            if record_header['record_size'] == (record_header['function'] >> 8) + 3:
                raise BadDataError('No bitmap embedded in metafile record')

            dib = s[length+L:record_header['record_size']*2]
            if verbose: print 'DIB', sformat(dib)

            # We have our DIB file! Almost done, but we need BMP header, which
            # requires image data offset. To find out the offset we will parse
            # DIB header. For now we just abort on all complex DIB files
            # (where image data doesn't go immediately after DIB header).
            _, dib_header = unwrap(dib, """I header_size == BITMAPINFOHEADER ?
                                           i width
                                           i height
                                           H planes
                                           H bit_count == BI_BITCOUNT_5 ?
                                           I compression
                                           I image_size
                                           i hres
                                           i vres
                                           I ncolors == 0 ?
                                           I nimpcolors""", 'DIB')
            if verbose: pprint(dib_header)

            # BMP header format
            BMP = '<2sIHHI'

            # now we know image data follows immediately after DIB header
            data_offset = struct.calcsize(BMP) + dib_header['header_size']
            file_size = struct.calcsize(BMP) + len(dib)
            bmp_header = struct.pack(BMP, 'BM', file_size, 0, 0, data_offset)
            yield bmp_header + dib

        s = s[record_header['record_size']*2:]

        if record_header['function'] == META_EOF and s:
            raise BadDataError("Metafile didn't end with end-of-file record")

def sformat(s):
    return '%d:%r' % (len(s), s if len(s) < 30 else (s[:40]+'...'+s[-20:]))

def unwrap(binary, spec, data_name=None):
    """Unwrap `binary` according to `spec`, return (consumed_length, data)

    Basically it's a convenient wrapper around struct.unpack. Each non-empty
    line in spec must be: <struct format> <field name> [<test> <action>]

    struct format - struct module format producing exactly one value
    field name - dictionary key to put unpacked value into
    test - optional test (unpacked) value should pass
    action - what to do if test failed: `!` (bad data) or `?` (unsupported)

    Example:
    >>> unwrap('\x0a\x00DATA\x00something else', '''h magic == 0x0a !
    ...                                             4s data''')
    (6, {'magic': 10, 'data': 'DATA'})
    """

    matches = [re.match("""(\w+)           # struct format
                           \s+
                           (\w+)           # field name
                           ((.+)\ ([!?]))? # optional test-action pair
                           $""", s.strip(), re.VERBOSE)
               for s in spec.split('\n') if s and not s.isspace()]

    for n, m in enumerate(matches):
        if not m: raise SyntaxError('Bad unwrap spec, LINE %d' % (n+1))

    fmt = '<' + ''.join(m.group(1) for m in matches)
    names = [m.group(2) for m in matches]
    tests = [(m.group(4), m.group(5)) for m in matches]

    # unpack binary data
    length = struct.calcsize(fmt)
    values = struct.unpack(fmt, binary[:length])

    # run optional tests
    for v, name, (test, action) in zip(values, names, tests):
        if test and not eval(name + test, {name: v}, globals()):
            adj = {'!': 'Bad', '?': 'Unsupported'}[action]
            raise BadDataError(' '.join(w for w in
                    [adj, data_name, name, '== %r' % v] if w))

    return length, dict(zip(names, values))

if __name__ == '__main__': # tests
    master = open('test/master.bmp').read()
    for filename in ['test/paintbrush', 'test/dib']:
        print '%s:' % filename
        olefield = open(filename, 'rb').read()
        for object_type, data in objects(olefield):
            print '\t- object %r %d bytes' % (object_type, len(data))
            if object_type == 'METAFILEPICT':
                for image in metafile_bmps(data):
                    assert master == image
                    print '\t\t* %d bytes BMP image: ok' % len(image)

    print 'test/paintbrush (higher-level API):'
    olefield = open('test/paintbrush', 'rb').read()
    for bmp in bmps(olefield):
        assert bmp.startswith(master)
        print '\t%d bytes BMP image: ok' % len(bmp)
