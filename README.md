OLE2Reader
==========

Python and Matlab code for importing data from OLE 2 files


Getting started
---------------
The code is based on the [Apache POI Java library](http://poi.apache.org). To get started, download a copy of the binary distribution [here](http://poi.apache.org/download.html). You'll need version 3.11-beta (or a later version) in order to read files that are larger than 2 GB.


Python 2.7
---------------

#### Required packages
In addition to the [Apache POI Java library](http://poi.apache.org), the Python code requires the Python packages [Numpy](http://www.numpy.org) and [JPype](https://github.com/originell/jpype) which must be installed separately. This can be done using [pip](https://github.com/pypa/pip):

```bash
pip install JPype1 numpy
```

#### Installation
```bash
pip install https://github.com/martinandersen/OLE2Reader/archive/master.zip
```

#### Code example
```python
# Import OLE2Reader and start Java Virtual Machine (JVM)
import OLE2Reader
OLE2Reader.startJVM(poifs_path = 'poi-library.jar')

# Create OLE2File object from OLE2 file
of = OLE2Reader.OLE2File('/path/to/my_ole2_file')

# List contents
of.ls('/')    

# Read document entry as a Numpy uint8 array
arr = of.readBytes('/path/to/entry')

# Read document entry as a Numpy uint16 array
arr = of.readShort('/path/to/entry')

# Read document entry as a Numpy uint32 array
arr = of.readInt('/path/to/entry')

# Read document entry as a Numpy float array
arr = of.readFloat('/path/to/entry')

# Read document entry as a string
s = of.readString('/path/to/entry')


# Close file
of.close()

# Shutdown JVM
OLE2Reader.shutdownJVM()
``` 


Matlab
---------------

#### Code example
```matlab
% Set Matlab path and Matlab Java path
javaaddpath('path/to/poi-library.jar')
javaaddpath('java/')   % path to ReadOLE2Entry.class
addpath('matlab/')     % path to OLE2Reader Matlab routines

% List contents
ole2_ls('my_ole2_file')

% Read document entry as byte array
arr = ole2_cat('my_ole2_file', '/path/to/entry', 'byte');

% Read document entry as short array
arr = ole2_cat('my_ole2_file', '/path/to/entry', 'short');

% Read document entry as int array
arr = ole2_cat('my_ole2_file', '/path/to/entry', 'int');

% Read document entry as float array
arr = ole2_cat('my_ole2_file', '/path/to/entry', 'float');

% Read document entry as string
arr = ole2_cat('my_ole2_file', '/path/to/entry', 'str');
```
