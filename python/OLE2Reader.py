"""
================
OLE2Reader
================

Python module for importing data from OLE 2 files.

License
----------------
OLE2Reader is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

OLE2Reader is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with Chompack.  If not, see <http://www.gnu.org/licenses/>.    


Requirements
----------------

This module depends on the `JPype <http://jpype.readthedocs.org>`_
Python extension, `Numpy <http://www.numpy.org>`_ and `Apache POI
<http://poi.apache.org/download.html>`_, a Java implementation of the
OLE 2 Compound Document format:  

  * Apache POI Java libray (version 3.11-beta or later to read 4GB+ files)
  * JPype Python extension (tested with version 0.5)

JPype can be installed using pip:

  $ pip install JPype1
  
  
Getting started
----------------

OLE2Reader uses JPype as a means to use the Apache POI to read data
from OLE 2 files. The following example illustrates how to:

  * start the Java virtual machine (with POI in the Java class path)
  * open a OLE 2 file
  * read a document/entry from the OLE 2 file system
  * close the OLE 2 file
  * shutdown the Java virtual machine 

```python
import OLE2Reader

OLE2Reader.startJVM(poifs_path = 'poi-library.jar')

of = OLE2Reader.OLE2File('/path/to/my_ole2_file')
of.ls()   # list contents 
arr = of.readInt('/path/to/entry')
of.close()

OLE2Reader.shutdownJVM()
``` 
"""

import jpype
import numpy

shutdownJVM = jpype.shutdownJVM

def startJVM(poifs_path, jvm_path = None):
    """Start JVM with POIFS in the Java class path."""
    if jvm_path is None: jvm_path = jpype.getDefaultJVMPath()
    jpype.startJVM(jvm_path,"-Djava.class.path="+poifs_path)
    return

class OLE2File(object):
    """OLE 2 file object

    The class relies on Apache's POI java archive (http://poi.apache.org)
    and JPype which allows Python programs full access to java class
    libraries.
    """

    def __init__(self, filename, buffer_size = None):
        """Opens an OLE 2 file system."""

        self._filename = filename
        File = jpype.JClass('java.io.File')
        self._fp = File(jpype.JString(filename))
        NPOIFSFileSystem = jpype.JPackage('org').apache.poi.poifs.filesystem.NPOIFSFileSystem
        self._fs = NPOIFSFileSystem(self._fp)
        self._root = self._fs.getRoot();
        self._DocumentInputStream = jpype.JPackage('org').apache.poi.poifs.filesystem.DocumentInputStream
        self._ByteBuffer = jpype.java.nio.ByteBuffer
        self._buffer_size = buffer_size
        if self._buffer_size is not None:
            self._bb = self._ByteBuffer.allocate(self._buffer_size)        
            self._buffer = self._bb.array()
        return

    def _getEntry(self, path):
        """Retrieves a handle to a DocumentEntry or a DirectoryEntry."""
        h = self._root
        if h is not None and path:
            pathstr = path.split('/');
            for p in pathstr:
                if p: h = h.getEntry(p)
        return h

    def ls(self, path = None):
        """
        Returns the directory contents as a dictionary where the keys
        are file and directory names. The values are either the file
        size in bytes or 'd' for directories.
        
        If no path is given, the contents of the root directory are returned.
        """
        h = self._getEntry(path)
        if not h: raise IOError
        if h.isDocumentEntry():
            entries = {h.getName() : h.getSize()}
        else:
            entries = {}
            for e in h.iterator(): # traverse document tree
                if e.isDocumentEntry():
                    entries[e.getName()] = e.getSize()
                else:
                    entries[e.getName()] = 'd'
        return entries

    def tree(self, path='/'):
        """
        Returns the directory contents as a dictionary where the keys
        are file and directory names. The values are either the file
        size in bytes or a dictionary for directories.
        
        If no path is given, the contents of the root directory are returned.
        """
        T = self.ls(path)
        stack = [(T,path)]
        while len(stack) > 0:
            Di,path = stack.pop()
            for key,e in Di.iteritems():
                if e == 'd':
                    Di[key] = self.ls(path+'/'+key)
                    stack.append((Di[key],path+'/'+key))
        return T

    def lstree(self, path='/'):
        """
        Prints a subtree.
        """
        T = self.tree(path)
        stack = sorted(zip(T.keys(),T.values()),reverse=True)
        indent = len(stack)*[(0,)]
        indent[0] = (indent[0][0],None)
        print("%s\n \\" % (path))
        while len(stack) > 0:
            k,v = stack.pop()
            t = indent.pop()
            if len(t) == 1:
                print("%s |- %s " % (t[0]*" ",k))
            else:
                print("%s `- %s " % (t[0]*" ",k))
            if type(v) is dict:
                print("%s \\ " % ((4+t[0])*" "))
                L = sorted(zip(v.keys(),v.values()),reverse=True)
                stack.extend(L)
                indent.extend(len(L)*[(t[0]+4,)])
                indent[-len(L)] = (indent[-len(L)][0],None)
        return
                
    def readBytes(self, entry):
        """
        Reads a document entry from the OLE 2 file system and returns
        the data as a Numpy array of bytes.
        """
        e = self._getEntry(entry)
        nbytes = e.getSize()
        if self._buffer_size is not None:
            if nbytes > self._buffer_size: raise BufferError
            buf = self._buffer
        else:
            bb = self._ByteBuffer.allocate(nbytes)        
            buf = bb.array()
            
        stream = self._DocumentInputStream(e)
        stream.readFully(buf, 0, nbytes);
        stream.close();

        java_str = jpype._jclass.JClass('java.lang.String')(buf, 0, nbytes, 'ISO-8859-1')        
        arr = numpy.array(numpy.frombuffer(buffer(java_str.toString()),dtype='uint16'),dtype='uint8')
        return arr

    def read(self, entry, dtype='uint8'):
        """
        Reads a document entry from the OLE 2 file system and returns
        the data as a Numpy array of type 'dtype'.
        """
        return numpy.frombuffer(buffer(self.readBytes(entry)),dtype=dtype)
    
    def readString(self, entry):
        """
        Reads a document entry from the OLE 2 file system and returns
        the data as a string.
        """
        return str(buffer(self.readBytes(entry)))

    def readShort(self, entry):
        """
        Reads a document entry from the OLE 2 file system and returns
        the data as a Numpy array of 16 bit unsigned integers.
        """
        return self.read(entry, '<u2')

    def readInt(self, entry):
        """
        Reads a document entry from the OLE 2 file system and returns
        the data as a Numpy array of 32 bit unsigned integers.
        """
        return self.read(entry, '<u4')

    def readFloat(self, entry):
        """
        Reads a document entry from the OLE 2 file system and returns
        the data as a Numpy array of 32 bit floats.
        """
        return self.read(entry, '<f4')

    def readDouble(self, entry):
        """
        Reads a document entry from the OLE 2 file system and returns
        the data as a Numpy array of 64 bit floats.
        """
        return self.read(entry, '<f8')
        
    def close(self):
        """
        Closes the OLE 2 file system.
        """
        self._fs.close()
        self._root = None
        return
    


