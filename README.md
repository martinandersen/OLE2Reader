OLE2Reader
==========

Python and Matlab code for importing data from OLE 2 files


Getting started
---------------
The code is based on the [Apache POI Java library](http://poi.apache.org). To get started, download a copy of the binary distribution [here](http://poi.apache.org/download.html). You'll need version 3.11-beta (or a later version) in order to read files that are larger than 4 GB.


Python
------

#### Required packages
The Python code requires the Python packages [Numpy](http://www.numpy.org) and [JPype](https://github.com/originell/jpype) which must be installed separately. This can be done using [pip](https://github.com/pypa/pip):

```bash
pip install JPype1 numpy --user
```

#### Code example
```python
import OLE2Reader

OLE2Reader.startJVM(poifs_path = 'poi-library.jar')

of = OLE2Reader.OLE2File('/path/to/my_ole2_file')
of.ls()   # list contents 
arr = of.readInt('/path/to/entry')
of.close()

OLE2Reader.shutdownJVM()
``` 


Matlab
------


