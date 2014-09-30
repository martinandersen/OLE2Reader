function d = ole2_cat(filename, entry, rtype)
% OLE2_CAT  Catenates entry in OLE 2 file
%
%    d = OLE2_CAT(filename, entry) catenates data from entry and
%    returns a byte array (little-endian). 
%
%    d = OLE2_CAT(filename,entry,rtype) catenates data from entry
%    and returns data as an array of type rtype (rtype can be 'byte',
%    'short', 'int', 'float, or 'str').
%
% The function relies on Apache's POI Java archive (http://poi.apache.org)
% and the Java class ReadOLE2Entry which must be available in
% Matlab's Java path. This can be done as follows:
%
%    javaaddpath('path/to/poi-library.jar')
%    javaaddpath('../java/')   % path to ReadOLE2Entry.class
%
% See also: javaaddpath, javaclasspath

% Copyright 2014 M. Andersen
% 
% OLE2Reader is free software; you can redistribute it and/or modify
% it under the terms of the GNU General Public License as published by
% the Free Software Foundation; either version 3 of the License, or
% (at your option) any later version.
%
% OLE2Reader is distributed in the hope that it will be useful,
% but WITHOUT ANY WARRANTY; without even the implied warranty of
% MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
% GNU General Public License for more details.
% 
% You should have received a copy of the GNU General Public License
% along with this program.  If not, see <http://www.gnu.org/licenses/>.

% return type
if nargin == 3
    assert(isstr(rtype),'rtype must be a string');
else
    rtype = 'byte';
end

% open file and find document handle
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import ReadOLE2Entry;
if strfind(filename, '~') == 1
    [filepath, name, ext] = fileparts(filename);
    % expand filepath to an absolute path
    filepath = cd(cd(filepath));
    filename = fullfile(filepath, [name ext]);
end
fp = java.io.File(filename);
fs = NPOIFSFileSystem(fp);
eptr = fs.getRoot();
if nargin >= 2
    pathstr = strsplit(entry,'/');
    for ii = 1:length(pathstr)
        if ~isempty(pathstr{ii}); eptr = eptr.getEntry(pathstr{ii}); end;
    end
end
assert(eptr.isDocumentEntry(),'Only document entries can be catenated.');

% read data
switch lower(rtype)
    case 'byte'
        d = typecast(ReadOLE2Entry.ReadBytes(eptr),'uint8')';
    case 'short'
        d = typecast(ReadOLE2Entry.ReadShort(eptr),'uint16')';
    case 'int'
        d = typecast(ReadOLE2Entry.ReadInt(eptr),'uint32')';
    case 'float'
        d = typecast(ReadOLE2Entry.ReadInt(eptr),'single')';        
    case 'single'
        d = typecast(ReadOLE2Entry.ReadInt(eptr),'single')';
    case 'double'
        d = typecast(ReadOLE2Entry.ReadDouble(eptr),'double')';
    case 'str'
        d = char(ReadOLE2Entry.ReadString(eptr));
    otherwise
        disp('rtype must be "byte", "short", "int", "float", "double", or "str"');
end

% clean up
fs.close()
