"""Extract files Microsoft Access OLE object fields"""

import struct
import re
import sys
from pprint import pprint

class BadDataError(Exception):
    pass

def parse_olefield(s, verbose=False):
    """Parse OLE object field and return iterator over 'objects' embedded

    Each iteration returns (object_type, object_content) tuple
    """

    # The best information about OLE objects in Microsoft Access fields:
    # * http://jvdveen.blogspot.com/2009/01/ole-and-accessing-files-embedded-in.html
    # * http://jvdveen.blogspot.com/2009/02/ole-and-accessing-files-embedded-in.html

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

        length, ole_header = unwrap(s, """I ole_version == 0x0501 ?
                                          I ole_format
                                          i object_type_len""")
        if verbose: pprint(ole_header)
        s = s[length:]

        length, ole_header_cont = unwrap(s, """{object_type_len}s object_type
                                               8s unknown
                                            """.format(**ole_header))
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

def parse_metafile(s, verbose=False):
    """Parse Metafile inside OLE field and return iterator over BMP files

    Can parse 'METAFILEPICT' objects from `parse_olefield`
    """

    # The content of METAFILEPICT object is a Windows Metafile,
    # but with 8 bytes prepended (don't know what they mean).
    #
    # Also see "Windows Metafile Format (wmf) Specification":
    # http://msdn.microsoft.com/en-us/library/cc215212.aspx

    # metafile header
    length, header = unwrap(s, """8s unknown
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
        # metafile record
        length, record_header = unwrap(s, """I record_size
                                             H function""")
        if verbose: pprint(record_header)

        if record_header['function'] == META_DIBSTRETCHBLT:
            # META_DIBSTRETCHBLT record function parameters
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

            # ensure record has bitmap (test copied from WMF spec, 2.3.1.3)
            if record_header['record_size'] == (record_header['function'] >> 8) + 3:
                raise BadDataError('No bitmap embedded in metafile record')

            dib = s[length+L:record_header['record_size']*2]
            if verbose: print 'DIB', sformat(dib)

            # We have our DIB file! Yes, but we have to cook BMP header,
            # which needs image data offset. To find out the offset we will
            # parse DIB. In this implementation we just abort on all complex
            # DIB files (where offset != BMP header size + DIB header size).

            # DIB header
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

def unwrap(binary, spec, data_name=''):
    """Unwrap `binary` according to `spec`, return (consumed_length, data)

    Basically it's a convenient wrapper around struct.unpack. Each non-empty
    line in spec must be: <struct format> <field name> [<test> <action>]

    struct format - struct module format producing exactly one value
    field name - dictionary key to put unpacked value into
    test - optional test (unpacked) value should pass
    action - what to do if test failed: `!` (bad data) or `?` (unsupported)

    Example:
    >>> unwrap('\x01\x00something else', '''h word
    ...                                     9s string == 'something' ?
                                         ''')
    (4, {'word': 1, 'string': 'something'})
    """

    matches = [re.match("""\s*
                           (\w+)           # struct format
                           \s+
                           (\w+)           # field name
                           ((.+)\ ([!?]))? # optional test-action pair
                           $
                        """, s, re.VERBOSE)
               for s in spec.split('\n') if s.strip()]

    for n, m in enumerate(matches):
        if not m: raise SyntaxError('Bad unwrap spec, LINE %d' % (n+1))

    fmt = '<' + ''.join(m.group(1) for m in matches)
    names = [m.group(2) for m in matches]
    tests = [(m.group(4), m.group(5)) for m in matches]

    length = struct.calcsize(fmt)
    fields = struct.unpack(fmt, binary[:length])

    if data_name: data_name += ' '
    for f, name, (test, action) in zip(fields, names, tests):
        if test and not eval(name + test, {name: f}, globals()):
            raise BadDataError('%s %s%s' %
                ('Bad' if action=='!' else 'Unsupported', data_name or '', name))

    return length, dict(zip(names, fields))

if __name__ == '__main__':
    paint = open('paintbrush_picture_big_boy').read()
    dib = open('dib_picture_big_boy').read()

    for olefield in (dib, paint):
        for object_type, data in parse_olefield(olefield):
            print object_type, sformat(data)
            if object_type == 'METAFILEPICT':
                for image in parse_metafile(data):
                    print 'image', sformat(image)
