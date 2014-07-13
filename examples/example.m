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

% Read document entry as double array
arr_double = ole2_cat('testfile.ole2', 'FloatingPointData/DoubleFile', 'double');

% Read document entry as string
s = ole2_cat('testfile.ole2', 'CharFile', 'str');