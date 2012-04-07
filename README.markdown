## olefield: parse Microsoft Access&reg; OLE object fields

`olefield` is a Python module for poor souls who need to extract data out of Microsoft Access&reg; database OLE object fields. Currently `olefield` knows how to extract BMP images.

Usage:

```python
>>> import olefield
>>> ole_content = '...' # you have to load oleobject field data somehow ;-)
>>> n = 1
>>> for bmp in olefield.bmps(ole_content):
...     open('%d.bmp' % n, 'wb').write(bmp)
...     n += 1
```

If the above fails or you need to extract something else (beside BMP), try lower level API:

```python
>>> for object_type, data in olefield.objects(ole_content):
...     print '%r %d bytes: %r' % object_type, len(data), data[:30]
```

Then send me an email telling me what you get ;-)

`olefield` doesn't yet support structured storage inside OLE fields. But it tries hard to skip data it doesn't understand. It should be easy to implement support for other files (pdf, doc).

I have only tested the module with Microsoft Access&reg; 2007.
