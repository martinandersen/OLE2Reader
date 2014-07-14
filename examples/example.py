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
