function varargout = ole2_ls(filename, d)
% OLE2_LS  Lists entries in OLE 2 file
%
%    OLE2_LS(filename) lists the entries in the root directory and
%    file sizes in bytes.
%
%    OLE2_LS(filename) lists the entries in the directory d.
%
%    D = OLE2_LS(filename) returns a struct array D with the entries
%    in the root directory. 
%
%    D = OLE2_LS(filename,d) returns a struct array D with the
%    entries in the directory d.
%
% The function relies on Apache's POI Java archive (http://poi.apache.org)
% which must be available in Matlab's Java path. This can be done as
% follows:
%
%    javaaddpath('path/to/poi-library.jar')
% 
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

% open file and find directory handle
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
fp = java.io.File(filename);
fs = NPOIFSFileSystem(fp);
dptr = fs.getRoot();
if nargin == 2
    pathstr = strsplit(d,'/');
    for ii = 1:length(pathstr)
        if ~isempty(pathstr{ii}); dptr = dptr.getEntry(pathstr{ii}); end;
    end
end
assert(dptr.isDirectoryEntry(),'Only directory entries can be listed.');

% traverse document tree
entries = get_entries(dptr);
D = cell(length(entries),2);
for k = 1:length(entries)
    entry = entries{k};
    D{k,1} = char(entry.getName());
    if entry.isDirectoryEntry()
        D{k,2} = 0;
    else        
        D{k,2} = entry.getSize();
    end
end
if nargout == 1
    varargout{1} = struct('name',D(:,1),'bytes',D(:,2));
elseif nargout == 0
    for k = 1:length(entries)
        if D{k,2} > 0
            fprintf(1,'%-32s%10i\n',D{k,1},D{k,2});
        else
           fprintf(1,'%-32s\n',D{k,1});
        end
    end    
end

% clean up
fs.close()

% --- auxiliary routines --- %
function entries = get_entries(dir_entry)
n = dir_entry.getEntryCount();
entries_iter = dir_entry.getEntries();
entries = cell(n,1);
for i = 1:n
    entries{i} = entries_iter.next();
end