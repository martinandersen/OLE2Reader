/*
  Copyright 2014 M. Andersen
  
  OLE2Reader is free software; you can redistribute it and/or modify
  it under the terms of the GNU General Public License as published by
  the Free Software Foundation; either version 3 of the License, or
  (at your option) any later version.
 
  OLE2Reader is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.
  
  You should have received a copy of the GNU General Public License
  along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/


/* 
  Compile with: 
     javac -cp $POI_JAR_FILE:. ReadOLE2Entry.java 
  or
     javac -target 1.6 -cp $POI_JAR_FILE:. ReadOLE2Entry.java 
*/

import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import java.lang.String;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.IntBuffer;
import java.nio.ShortBuffer;
import java.nio.ByteOrder;

public class ReadOLE2Entry {

    public static double[] ReadDouble(DocumentEntry document) throws IOException {
	DocumentInputStream stream = new DocumentInputStream(document);
	int len = document.getSize();
	byte[] buf = new byte[len];
	double[] bufDbl = new double[len/8];
	try {
	    stream.readFully(buf, 0, len);
	} finally {
	    stream.close();
	}
	ByteBuffer.wrap(buf).order(ByteOrder.LITTLE_ENDIAN).asDoubleBuffer().get(bufDbl);
	return bufDbl;
    }

    public static int[] ReadInt(DocumentEntry document) throws IOException {
	DocumentInputStream stream = new DocumentInputStream(document);
	int len = document.getSize();
	byte[] buf = new byte[len];
	int[] bufInt = new int[len/4];
	try {
	    stream.readFully(buf, 0, len);
	} finally {
	    stream.close();
	}
	ByteBuffer.wrap(buf).order(ByteOrder.LITTLE_ENDIAN).asIntBuffer().get(bufInt);
	return bufInt;
    }

    public static short[] ReadShort(DocumentEntry document) throws IOException {
	DocumentInputStream stream = new DocumentInputStream(document);
	int len = document.getSize();
	byte[] buf = new byte[len];
	short[] bufShort = new short[len/2];
	try {
	    stream.readFully(buf, 0, len);
	} finally {
	    stream.close();
	}
	ByteBuffer.wrap(buf).order(ByteOrder.LITTLE_ENDIAN).asShortBuffer().get(bufShort);
	return bufShort;
    }

    public static byte[] ReadBytes(DocumentEntry document) throws IOException {
	DocumentInputStream stream = new DocumentInputStream(document);
	int len = document.getSize();
	byte[] buf = new byte[len];
	try {
	    stream.readFully(buf, 0, len);
	} finally {
	    stream.close();
	}
	return buf;
    }

    public static String ReadString(DocumentEntry document) throws IOException {
	DocumentInputStream stream = new DocumentInputStream(document);
	int len = document.getSize();
	byte[] buf = new byte[len];
	try {
	    stream.readFully(buf, 0, len);
	} finally {
	    stream.close();
	}
	String str = new String(buf);
	return str;
    }
}
