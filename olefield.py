"""Extract files Microsoft Access OLE object fields"""

import struct
import sys
from pprint import pprint


def sformat(s):
    return '%d:%r' % (len(s), s if len(s) < 30 else (s[:40]+'...'+s[-20:]))

def unwrap(s, spec):
    """Return (consumed_length, data_dict)"""
    fmt = '<'
    names = []
    for spec in spec.split('\n'):
        if spec:
            t, _, name = spec.strip().partition(' ')
            t = t.strip()
            name = name.strip()
            fmt += t
            names.append(name)

    length = struct.calcsize(fmt)

    return length, dict(zip(names, struct.unpack(fmt, s[:length])))

class BadDataError(Exception):
    pass

def parse_olefield(s, verbose=False):
    """Parse OLE object field and return iterator over 'objects' embedded

    Each iteration returns (object_type, object_content) tuple
    """

    # The most helpfull resource about OLE objects in Microsoft Access fields:
    # * http://jvdveen.blogspot.com/2009/01/ole-and-accessing-files-embedded-in.html
    # * http://jvdveen.blogspot.com/2009/02/ole-and-accessing-files-embedded-in.html

    length, header = unwrap(s, """h signature
                                  h header_size
                                  i object_type
                                  h friendly_len
                                  h class_len
                                  h friendly_off
                                  h class_off""")
    if verbose: pprint(header)
    if header['signature'] != struct.unpack('<h', '\x15\x1c')[0]:
        raise BadDataError('Bad signature')

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
                                          3s footer""")
            if verbose: print 'footer %r' % footer
            if footer['footer'] != '\xad\x05\xfe':
                raise BadDataError('Bad footer')
            break

        length, ole_header = unwrap(s, """I ole_version
                                          I ole_format
                                          i object_type_len""")
        if verbose: pprint(ole_header)
        if ole_header['ole_version'] != struct.unpack('<I', '\x01\x05\x00\x00')[0]:
            raise BadDataError('Unsupported OLE version')
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
        s = s[length+block_header['data_block_len']:]
        yield ole_header_cont['object_type'].rstrip('\x00'), data

META_EOF = 0x0000
META_DIBSTRETCHBLT = 0x0b41
BITMAPINFOHEADER = 40
BI_BITCOUNT_5 = 0x0018

def parse_metafile(s, verbose=False):
    """Parse Metafile inside OLE field and return iterator over BMP files

    suitable for 'METAFILEPICT' OLE field object `parse_olefield` return
    """

    # The content of METAFILEPICT object is a Windows Metafile,
    # but with 8 bytes prepended (don't know what they mean).
    #
    # Metafile spec: "Windows Metafile Format (wmf) Specification",
    # http://download.microsoft.com/download/5/0/1/501ED102-E53F-4CE0-AA6B-B0F93629DDC6/WindowsMetafileFormat(wmf)Specification.pdf

    # metafile header
    length, header = unwrap(s, """8s unknown
                                  H type
                                  H header_size
                                  H version
                                  I metafile_size
                                  H num_of_objects
                                  I max_record_len
                                  H unused_should_be_0""")
    if verbose: pprint(header)
    if header['type'] != struct.unpack('<h', '\x01\x00')[0]:
        raise BadDataError('Unknown metafile type')
    if header['version'] != struct.unpack('<h', '\x00\x03')[0]:
        raise BadDataError('Unsupported metafile version')
    if header['num_of_objects'] > 0:
        raise BadDataError('Metafile with graphics objects not supported')

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
            _, dib_header = unwrap(dib, """I header_size
                                           i width
                                           i height
                                           H planes
                                           H bit_count
                                           I compression
                                           I image_size
                                           i hres
                                           i vres
                                           I ncolors
                                           I nimpcolors""")
            if verbose: pprint(dib_header)

            if dib_header['header_size'] != BITMAPINFOHEADER:
                raise BadDataError('Unsupported DIB header type')
            if dib_header['bit_count'] != BI_BITCOUNT_5:
                raise BadDataError('Unsupported DIB bit_count value')
            if dib_header['ncolors'] != 0:
                raise BadDataError('Unsupported DIB ncolors value')

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

if __name__ == '__main__':
    paint = open('paintbrush_picture_big_boy').read()
    dib = open('dib_picture_big_boy').read()

    for olefield in (dib, paint):
        for object_type, data in parse_olefield(olefield):
            print object_type, sformat(data)
            if object_type == 'METAFILEPICT':
                for image in parse_metafile(data):
                    print 'image', sformat(image)
