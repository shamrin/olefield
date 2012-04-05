## olefield: parse Microsoft Access&reg; OLE object fields

`olefield` is a Python module for poor souls who need to extract data out of Microsoft Access&reg; OLE object fields.

Currently `olefield` can only extract BMP images and doesn't support structured storage inside OLE fields. But it tries hard to skip data it doesn't understand. It should be easy to implement support for other files (pdf, doc).

I have only tested the module with Microsoft Access&reg; 2007.

Usage:

```python
>>> from olefield import parse_olefield, parse_metafile
>>> ole_content = '...' # you have to load oleobject field content somehow ;-)
>>> n = 1
>>> for object_type, data in parse_olefield(ole_content):
...     if object_type == 'METAFILEPICT':
...         for image in parse_metafile(data):
...             open('%d.bmp' % n, 'wb').write(image)
...             n += 1
```
