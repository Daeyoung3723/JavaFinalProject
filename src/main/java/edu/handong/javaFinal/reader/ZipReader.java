package edu.handong.javaFinal.reader;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Enumeration;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipFile;


public class ZipReader {

	public ArrayList<String> readFileInZip(String path, String path1, int order) {
		ZipFile zipFile;
		try {
			zipFile = new ZipFile(path + "/" + path1 + ".zip");
			
			Enumeration<? extends ZipArchiveEntry> entries = zipFile.getEntries();

		    if(order == 1) {
		    	ZipArchiveEntry entry = entries.nextElement();
		        InputStream stream = zipFile.getInputStream(entry);
		    
		        ExcelReader myReader = new ExcelReader();
		        
		        return myReader.getData(stream);
		    } else if(order == 2) {
		    	ZipArchiveEntry entry = entries.nextElement();
		    	entry = entries.nextElement();
		        InputStream stream = zipFile.getInputStream(entry);
		    
		        ExcelReader myReader = new ExcelReader();
		        
		        return myReader.getData(stream);
		    	
		    }
		    
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return null;
	}
}

