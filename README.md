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
OLE2Reader.startJVM(poifs_path = "path/to/java/poi-3.xx.jar") 

of = OLE2Reader.OLE2File('testfile.ole2')

# Print directory tree
of.lstree()

# Read string
s = of.readString('CharFile')

# Read document entry as a Numpy uint8 array
arr_char = of.readBytes('ByteFile')

# Read document entry as a Numpy uint16 array
arr_short = of.readShort('ShortFile')

# Read document entry as a Numpy uint32 array
arr_int = of.readInt('IntFile')

# Read document entry as a Numpy float array
arr_float = of.readFloat('FloatingPointData/FloatFile')

# Read document entry as a Numpy float array
arr_double = of.readDouble('FloatingPointData/DoubleFile')

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
javaaddpath('path/to/java/poi-3.xx.jar')
javaaddpath('../java/')   % path to ReadOLE2Entry.class
addpath('../matlab/')     % path to OLE2Reader Matlab routines

% List contents
ole2_ls('testfile.ole2')

% Read document entry as byte array
arr_byte = ole2_cat('testfile.ole2', 'ByteFile', 'byte');

% Read document entry as short array
arr_short = ole2_cat('testfile.ole2', 'ShortFile', 'short');

% Read document entry as int array
arr_int = ole2_cat('testfile.ole2', 'IntFile', 'int');

% Read document entry as float array
arr_float = ole2_cat('testfile.ole2', 'FloatingPointData/FloatFile', 'float');

% Read document entry as string
s = ole2_cat('testfile.ole2', 'CharFile', 'str');
```
