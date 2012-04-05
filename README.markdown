## olefield: parse Microsoft Access&reg; OLE object fields

`olefield` is a Python module for poor souls who need to extract data out of Microsoft Access&reg; OLE object fields.

Usage:

```python
>>> from olefield import parse_olefield, parse_metafile
>>> ole_content = '<content of your ole field>'
>>> n = 1
>>> for object_type, data in parse_olefield(ole_content):
...     if object_type == 'METAFILEPICT':
...         for image in parse_metafile(data):
...             open('%d.bmp' % n, 'wb').write(image)
...             n += 1
```
