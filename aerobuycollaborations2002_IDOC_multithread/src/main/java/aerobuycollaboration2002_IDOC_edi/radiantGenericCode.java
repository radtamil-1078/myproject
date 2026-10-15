
package aerobuycollaboration2002_IDOC_edi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;


public class radiantGenericCode {
	static Properties  prop = new Properties();
	static FileInputStream ip;
	static FileOutputStream out;
	public static void mkdir(String path) throws Exception {
		File dir = new File(path);
		if (!dir.exists()) {
			dir.mkdir();
			dir.setWritable(true, false);
			dir.setReadable(true, false);
			dir.setExecutable(true, false);
		}
	}	
 public static String getDateTime(String dateType,String dateFormat,int daysCount) throws Exception {
		 
		 SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);
		 Calendar c;
		 String dateValue;
		 if(dateType.equals("currentDate")) {
			 Date date = new Date();
	         dateValue = sdf.format(date);
		 }else if(dateType.equals("futureDate")){
			  c = Calendar.getInstance();
			    c.add(Calendar.DAY_OF_MONTH, daysCount);  
			    dateValue = sdf.format(c.getTime()); 
		 }else {
			 c = Calendar.getInstance();
			 c.add(Calendar.DAY_OF_MONTH,- daysCount);  
			    dateValue = sdf.format(c.getTime()); 
		 }        
        return dateValue;

    }
	public static String getFileNameInAFolder(String dir) throws Exception {
		 String fileName = null;
		File file = new File(dir);
       String[] fileList = file.list();
       for(String name:fileList){
       	fileName= name;    
       }
		return fileName;		
	}
	public static String moveFile(String source,String destination) {
		File a = new File(source);
		a.renameTo(new File(destination));
		a.delete();
		return null;
	}
	public static void zip(String source,String destination) throws Exception{
		try (ZipOutputStream zipOut = new ZipOutputStream(new FileOutputStream(destination))) {
		    File fileToZip = new File(source);
		    zipOut.putNextEntry(new ZipEntry(fileToZip.getName()));
		    Files.copy(fileToZip.toPath(), zipOut);
		}
	}
	public static void unzip(String source,String destination) throws Exception{		
	    String password = "";
	    try {
	         ZipFile zipFile = new ZipFile(source);
	         if (zipFile.isEncrypted()) {
	            zipFile.setPassword(password);
	         }
	         zipFile.extractAll(destination);
	    } catch (ZipException e) {
	        e.printStackTrace();
	    }
	}
	public static String generateRandomString(int length, String mode, String specialCharatersValue) throws Exception {
        StringBuffer buffer = new StringBuffer();
        String characters = "";
        switch (mode) {
        case "Alpha":
            characters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
            break;
        case "Numeric":
            characters = "1234567890";
            break;
        case "AlphaNumeric":
            characters = "1abcdefghijklmnop2345qrstuvwxyzABCDEFGHI6JKLMNOPQRS78TUVWXYZ90";
            break;
        case "Special characters":
            if (specialCharatersValue.length() > 0) {
                characters = specialCharatersValue;
            } else {
                characters = "~`@#$%^&*()_-+={}[]|\"':;/?>.<,";
            }
            break;
        }
        int charactersLength = characters.length();

        for (int i = 0; i < length; i++) {
            double index = Math.random() * charactersLength;
            buffer.append(characters.charAt((int) index));
        }
        String returnValue=buffer.toString();
        if(returnValue.equals("0")){
            returnValue="1";
        }
        return returnValue;
    }
	public static String readPropertyFile(String propFilePath,String key) throws Exception {
		 
		  ip = new FileInputStream(propFilePath);
		  prop.load(ip);
		  String value=prop.getProperty(key);		  
	      ip.close();	        
		return value;		
	}
	public static String writePropertyFile(String propFilePath,String key,String value) throws Exception {
		ip = new FileInputStream(propFilePath);    	
    	prop.load(ip);
    	ip.close();
    	out = new FileOutputStream(propFilePath);
        prop.setProperty(key,value);
    	prop.store(out, null);
    	out.close();	        
		return null;		
	}
}
