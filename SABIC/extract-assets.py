"""Extract original cover artwork without rasterizing the PDF's text.

Requires pypdf: python3 -m pip install pypdf
"""
from pathlib import Path
import struct
import zlib

from pypdf import PdfReader


def png(image):
    width, height = image['/Width'], image['/Height']
    assert image['/ColorSpace'] == '/DeviceRGB' and image['/BitsPerComponent'] == 8
    rgb = image.get_data()
    mask = image.get('/SMask')
    alpha = mask.get_object().get_data() if mask else None
    channels = 4 if alpha else 3
    assert len(rgb) == width * height * 3
    if alpha:
        assert len(alpha) == width * height
        pixels = bytearray(width * height * 4)
        pixels[0::4], pixels[1::4], pixels[2::4] = rgb[0::3], rgb[1::3], rgb[2::3]
        pixels[3::4] = alpha
    else:
        pixels = rgb
    rows = b''.join(b'\0' + pixels[y * width * channels:(y + 1) * width * channels]
                    for y in range(height))

    def chunk(kind, data):
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))

    return (b'\x89PNG\r\n\x1a\n'
            + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 6 if alpha else 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(rows)) + chunk(b'IEND', b''))


if __name__ == '__main__':
    folder = Path(__file__).resolve().parent
    reader = PdfReader(folder / 'Template.pdf')
    for name, object_id in [('logo-yellow', 22), ('ribbon-yellow', 24), ('ribbon-blue', 26)]:
        (folder / f'{name}.png').write_bytes(png(reader.get_object(object_id)))
    print('Extracted original SABIC logo and two ribbon assets.')
