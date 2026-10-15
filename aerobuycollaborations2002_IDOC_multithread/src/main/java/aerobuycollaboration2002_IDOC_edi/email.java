package aerobuycollaboration2002_IDOC_edi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Properties;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.net.ssl.SSLContext;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class email {
	 private static String homeDir="/home/automation/jenkins/aeroxchange/output/aerobuy/output/";
	// private static String homeDir="D:/home/automation/jenkins/aeroxchange/output/aerorepair/output/";
	 private static ArrayList<String> outboundGeneratedSuccessfully;
	 private static ArrayList<String> posted;
	 private static ArrayList<String> passTestCase;
	 private static ArrayList<String> failTestCase;
	 
	
	 public static String emailTriggerGmail(String functionality, String expectationPassCount, String expectationFailCount ) throws Exception {
		Thread.sleep(400000);
		 Properties props = new Properties();
			props.put("mail.smtp.auth", "true");
			props.put("mail.smtp.starttls.enable", "true");
			props.put("mail.smtp.host", "smtp.gmail.com");
			props.put("mail.smtp.port", "587");
			props.put("mail.debug", "false");
			SSLContext context = SSLContext.getInstance("TLS");
			context.init(null, null, null);
			String[] supportedProtocols = context.getDefaultSSLParameters().getProtocols();
			System.out.println(Arrays.toString(supportedProtocols)+"Protocol==");
		    Session session = Session.getInstance(props, new Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication("aeroxchanges@gmail.com", "xfifvqhucrmsnpum");
			}
		});
		try {
			outboundGeneratedSuccessfully = new ArrayList<String>();
			posted = new ArrayList<String>();
			passTestCase = new ArrayList<String>();
			failTestCase = new ArrayList<String>();
			DateFormat dateFormat1 = new SimpleDateFormat("MMddyyyy");
			Date CUR1 = new Date();
			String filename = homeDir + dateFormat1.format(CUR1) + "/edib2b/Aerorepair.xlsx";
	//		String filename = excelPath + dateFormat1.format(CUR1) + "/edib2b/"+functionality+".xls";
			outputExcel(filename,functionality);
			int outboundSuccessCount = outboundGeneratedSuccessfully.size();
			int postedCount = posted.size();
//			int passCount = passTestCase.size();
//			int failCount = failTestCase.size();
			MimeMessage msg = new MimeMessage(session);
			msg.setFrom(new InternetAddress("aeroxchanges@gmail.com"));
	 //       msg.setRecipients(Message.RecipientType.TO, "automationtesting@aeroxchange.com"+","+"prasanth.srinivasan@radiantglobaltech.com"+","+"maria.sebastin@radiantglobaltech.com"+","+"naveenkumar.m@radiantglobaltech.com"+","+"tamilselvan.s@radiantglobaltech.com"+","+"santhosh.raja@radiantglobaltech.com"+","+"naveen.venkatesh@radiantglobaltech.com");
	        msg.setRecipients(Message.RecipientType.TO, "tamilselvan.s@radiantglobaltech.com");
			msg.setSubject(functionality);
			msg.setSentDate(new Date());
			Multipart multipart = new MimeMultipart();
			MimeBodyPart htmlPart = new MimeBodyPart();
			String htmlContent = "<html><head><style>table, td,th {border: 1px solid black;}</style></head><body>" 
		        		+ "<h1>Result</h1>"
	       		+"<table><tr><th>Sl.No</th><th>Functionality</th><th>Scenarios</th><th>Expectation</th><th>Actual</th></tr><tr><td>1</td><td>Positive</td><td>Pass</td><td>"+expectationPassCount+"</td><td>"+outboundSuccessCount+"</td></tr><tr><td>2</td><td>Negative</td><td>Fail</td><td>"+expectationFailCount+"</td><td>"+postedCount+"</td></tr></table>"
		        		+ "</body>"
		        		+ "</html>";
			htmlPart.setContent(htmlContent, "text/html");
			multipart.addBodyPart(htmlPart);
			MimeBodyPart attachementPart = new MimeBodyPart();
			attachementPart.attachFile(new File(filename));
			multipart.addBodyPart(attachementPart);
			msg.setContent(multipart);
			Transport.send(msg);
			System.out.println("Outbound is Generated Sucessfully");
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return null;

	}
	public static String emailTriggerProton(String functionality,String expectationPassCount,String expectationFailCount) throws Exception {		 
	        Thread.sleep(10000);       
	        Properties props = new Properties();
	        props.put("mail.smtp.auth", "true");
	        props.put("mail.smtp.host", "127.0.0.1");	      
	        props.put("mail.smtp.port", "1026");
	        props.put("mail.smtp.starttls.enable", "false");
	        props.put("mail.smtp.ssl.enable", "false");	 
	        Session session = Session.getInstance(props, new Authenticator() {
	            protected PasswordAuthentication getPasswordAuthentication() {
	               // return new PasswordAuthentication("aex.preview.auto02@pm.me","m4B0knnZuAzC2YJjDXYiIw");  // Proton Bridge password Local
	                return new PasswordAuthentication("aex.preview.auto02.01@pm.me","46I4h3PplnV0sEFLv0jU0A"); // Proton Bridge password AWS
	            }
	        });
	        try {
	            outboundGeneratedSuccessfully = new ArrayList<>();
	            posted = new ArrayList<>();
	            passTestCase = new ArrayList<>();
	            failTestCase = new ArrayList<>();	 
	            DateFormat dateFormat1 = new SimpleDateFormat("MMddyyyy");
	            Date CUR1 = new Date();	 
	            String filename = homeDir + dateFormat1.format(CUR1) + "/2002_idoc/AeroBuy.xlsx";
	            outputExcel(filename, functionality);	 
	            int outboundSuccessCount = outboundGeneratedSuccessfully.size();
	            int postedCount = posted.size();	 
	            MimeMessage msg = new MimeMessage(session);	 
	            msg.setFrom(new InternetAddress("aex.preview.auto02.01@pm.me"));	 
	             msg.setRecipients(Message.RecipientType.TO, "tamilselvan.s@radiantglobaltech.com");
	            msg.setSubject(functionality);
	            msg.setSentDate(new Date());
	            Multipart multipart = new MimeMultipart();           
	            MimeBodyPart htmlPart = new MimeBodyPart();
	            String htmlContent = "<html><head><style>table, td,th {border: 1px solid black;}</style></head><body>"
		        		+ "<h1>Result</h1>"
	       		+"<table><tr><th>Sl.No</th><th>Functionality</th><th>Scenarios</th><th>Expectation</th><th>Actual</th></tr><tr><td>1</td><td>Positive</td><td>Pass</td><td>"+expectationPassCount+"</td><td>"+outboundSuccessCount+"</td></tr><tr><td>2</td><td>Negative</td><td>Fail</td><td>"+expectationFailCount+"</td><td>"+postedCount+"</td></tr></table>"
		        		+ "</body>"
	        		+ "</html>";
			htmlPart.setContent(htmlContent, "text/html");
			multipart.addBodyPart(htmlPart);
			MimeBodyPart attachementPart = new MimeBodyPart();
			attachementPart.attachFile(new File(filename));
			multipart.addBodyPart(attachementPart);
			msg.setContent(multipart);
			Transport.send(msg);
			System.out.println("Outbound is Generated Sucessfully");
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return null;
	    }

	public static void outputExcel(String filename, String functionality ) throws Exception {
		FileInputStream output = new FileInputStream(new File(filename));
		XSSFWorkbook wb = new XSSFWorkbook(output);
		XSSFSheet worksheet = wb.getSheetAt(0);
		int rowCount = worksheet.getPhysicalNumberOfRows();
		for (int i = 1; i < rowCount; i++) {
			String actualResult = worksheet.getRow(i).getCell(14).getStringCellValue();
			if (actualResult.equals("Pass")) {
				outboundGeneratedSuccessfully.add(actualResult);
			} else if (actualResult.equals("Fail")) {
				posted.add(actualResult);
			}
			String testCase = worksheet.getRow(i).getCell(15).getStringCellValue();
			if (testCase.equals("Pass")) {
				passTestCase.add(testCase);
			} else if (testCase.equals("Fail")) {
				failTestCase.add(testCase);
			}
		}
		output.close();
		FileOutputStream output_file = new FileOutputStream(new File(filename));
		wb.write(output_file);
		output_file.close();
		wb.close();
	}
}
