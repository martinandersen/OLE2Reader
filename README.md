OLE2Reader
==========
* * *
*This project is no longer maintained. Python users may use the [olefile](https://github.com/decalage2/olefile) package instead.*
* * *

Python and Matlab code for importing data from OLE 2 files


Getting started
---------------
The code is based on the [Apache POI Java library](http://poi.apache.org). To get started, download a copy of the binary distribution [here](http://poi.apache.org/download.html). You'll need version 3.11-beta (or a later version) in order to read files that are larger than 2 GB. After downloading the jar file, we recommend setting the environment variable `POI_JAR_FILE` to the full path of the jar file:

```bash
export POI_JAR_FILE=/path/to/apache/poi-3.xx.jar
```


Python 2.7
---------------

#### Required packages
In addition to the [Apache POI Java library](http://poi.apache.org), the Python code requires the Python packages [Numpy](http://www.numpy.org) and [JPype](https://github.com/originell/jpype) which must be installed separately. This can be done using [pip](https://github.com/pypa/pip):

```bash
pip install numpy Jpype1
```

#### Installation

```bash
pip install https://github.com/martinandersen/OLE2Reader/archive/master.zip
```

#### Code example

```python
import OLE2Reader

# Start JVM with Apache POI jar file in Java path

# Method I: The environment variable POI_JAR_FILE contains 
# the full path to the Apache POI jar file 
print(OLE2Reader.POI_JAR_FILE)
OLE2Reader.startJVM()  

# Method II:
# OLE2Reader.start_jvm(poi_jar_file = 'path/to/apache/poi-3.xx.jar')

of = OLE2Reader.OLE2File('testfile.ole2')

# Print directory tree
of.lstree()

# Read string
s = of.readString('CharFile')

# Read document entry as a Numpy uint8 array
arr_byte = of.readBytes('ByteFile')

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
javaaddpath(getenv('POI_JAR_FILE'))
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
