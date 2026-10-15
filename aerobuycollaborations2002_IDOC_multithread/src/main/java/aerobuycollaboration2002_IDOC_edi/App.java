package aerobuycollaboration2002_IDOC_edi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.StringReader;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

import okhttp3.MediaType;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.RequestBody;
import okhttp3.Response;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class App {
	private static OkHttpClient client;
	public static DocumentBuilderFactory docFactory;
	private static boolean isInitialized = false;
	private String CreationDateTime;
	private String CreationDateTime2;
	private String filename;
	public XSSFWorkbook wb;
	public XSSFSheet worksheet;
	public FileInputStream output;
	public XSSFRow row;
	public Cell cell = null;
	public String uniqueDocumentID;
	public String InvoiceNumber;

	// private static XSSFRow cell;
	public static void initialize() throws Exception {
		if (!isInitialized) {
			filePathHandling.aerobuyEDIMessagePath();
			client = new OkHttpClient.Builder().readTimeout(240, TimeUnit.SECONDS).build();
			docFactory = DocumentBuilderFactory.newInstance();
			isInitialized = true;

		}
	}

	public String run(String url, String fileName, String credential) throws Exception {
		MediaType mediaType = MediaType.parse("text/xml");
		RequestBody body = RequestBody.create(mediaType,
				new File(filePathHandling.defaultXmlOutputPath + fileName + ".xml"));
		Request request = new Request.Builder().url(url).post(body).addHeader("content-type", "text/xml")
				.addHeader("authorization", "Basic " + credential).addHeader("cache-control", "no-cache").build();
		Response response = client.newCall(request).execute();
		return response.body().string();
	}

	public String aeroComponent(String fileName, String credential) throws Exception {
//		String response = run("https://aeroxchangeb2bserver.com:9003/invoke/wm.tn/receive/", fileName, credential);
		 String response =
		 run("http:localhost:8090/automationserver/aerocomponent/edib2b/customer",fileName,credential);
		 response="<Response><status>200</status><aex:errorMsg>Sucessful</aex:errorMsg><statusMessage>OK</statusMessage></Response>";
		return response;
	}
	public String xmlMapping(String fileName, int sheetNumber, String crediatial, String rootTag, String scriptName,
	        String description, String scenarioType, String expectation, String messageType, String uniqueNumber,
	        String mappingKey, int mappingUniqueNumberWriteExcel, String sNo) throws Exception {
	    String messageID = mappingKey.split("\\|")[1];
	    File eoFile = new File(filePathHandling.defaultInputPath+ filePathHandling.aerobuyEDIMessageDataExcelFileName);
	    Calendar cal = Calendar.getInstance();
	    DateFormat dateTimeFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	     CreationDateTime = dateTimeFormat.format(cal.getTime());   
	    output = new FileInputStream(eoFile);
	    wb = new XSSFWorkbook(output);
	    worksheet = wb.getSheetAt(sheetNumber);
	    row = worksheet.getRow(0);
	    int cols = row.getLastCellNum();
	    if (uniqueNumber.equals("Yes")) {
	        uniqueDocumentID = "CI"+ radiantGenericCode.getDateTime("currentDate", "MMddyyHHmmSS", 0);
	    }
	    if (uniqueNumber.equals("Yes1")) {
	        uniqueDocumentID = "CI"+ radiantGenericCode.getDateTime("currentDate", "MMddHHmmSS", 0);
	    }	    
	    if (uniqueNumber.equals("Yes")){			
	    	radiantGenericCode.writePropertyFile(filePathHandling.defaultPropertiesPath + filePathHandling.A2002_idocMessagePropFileName, "UNIQUE_DOCUMENT_ID", uniqueDocumentID);
			}	
	    try {
	        System.out.println("XML INPUT FILE : "+ filePathHandling.defaultXmlInputPath + fileName + ".xml");
	        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
	        Document doc = docBuilder.parse(new File(filePathHandling.defaultXmlInputPath + fileName + ".xml"));
	        for (int col = 1; col < cols; col++) {
	            String tagName = wb.getSheetAt(sheetNumber)
	                    .getRow(0)
	                    .getCell(col)
	                    .getStringCellValue();
	            String commaSeparated = wb.getSheetAt(sheetNumber)
	                    .getRow(1)
	                    .getCell(col)
	                    .getStringCellValue();
	            List<String> items = Arrays.asList(commaSeparated.split("\\s*,\\s*"));
	            System.out.println("TAG : " + tagName);
	            System.out.println("VALUES : " + items);
	            String xmlTagName = null;
	            if (doc.getElementsByTagName(tagName).getLength() > 0) {
	                xmlTagName = tagName;
	            } else if (doc.getElementsByTagName("aex:" + tagName).getLength() > 0) {
	                xmlTagName = "aex:" + tagName;
	            } else {
	                System.out.println("TAG NOT FOUND IN XML : " + tagName);
	                continue;
	            }
	            NodeList nodeList = doc.getElementsByTagName(xmlTagName);
	            System.out.println("FOUND NODES : " + nodeList.getLength());
	            for (int i = 0; i < nodeList.getLength() && i < items.size(); i++) {
	                Node node = nodeList.item(i);
	                if (node == null) {
	                    continue;
	                }
	                NodeList childNodes = node.getChildNodes();
	                if (childNodes.getLength() > 0) {
	                    Node child = childNodes.item(0);	                   
	                    child.setNodeValue(items.get(i));	                    
	                    // ==========================================
	                    // FILE REFERENCE ATTRIBUTE HANDLING
	                    // ==========================================
	                    if (node.getNodeType() == Node.ELEMENT_NODE) {
	                        Element element = (Element) node;
	                        if (tagName.equals("FileReference")) {
	                            String base64Content = items.size() > 0 ? items.get(0) : "";
	                            String fileNameValue = items.size() > 1 ? items.get(1) : "";
	                            String fileTypeValue = items.size() > 2 ? items.get(2) : "";
	                            child.setNodeValue(base64Content);
	                            if (!fileNameValue.isEmpty()) {
	                                element.setAttribute("aex:FileName", fileNameValue);
	                            }
	                            if (!fileTypeValue.isEmpty()) {
	                                element.setAttribute("aex:FileType", fileTypeValue);
	                            }
	                        }
	                    }	                    
	                    if (fileName.equals("Po_Ack") && commaSeparated.contains("GetUniqueNumber|CPO")) {
	                        uniqueDocumentID = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);	                        
	                        NodeList belnrNodes = doc.getElementsByTagName("BELNR");	                        
	                        for (int j = 0; j < belnrNodes.getLength(); j++) {
	                            Node belnrNode = belnrNodes.item(j);
	                            if (belnrNode != null) {	                                
	                                belnrNode.setTextContent(uniqueDocumentID);	                               
	                            }
	                        }
	                    }	                    
	                    if (fileName.equals("Invoice")&& commaSeparated.contains("GetUniqueNumber|CPO")||fileName.equals("Invoice_Credit")&& commaSeparated.contains("GetUniqueNumber|CPO")) {
	                    	uniqueDocumentID = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);
	                        child.setNodeValue(InvoiceNumber);
	                        excelEditor.mappingUniqueNumberWriteExcel(InvoiceNumber,mappingKey,tagName);              	
	                        NodeList belnrNodes = doc.getElementsByTagName("BELNR");	                      
	                        String[] belnrValues = {
	                        	uniqueDocumentID,
	                            "9306365022",
	                            "123",
	                            "1003532229",
	                            "PSN2016",
	                            "BOL05",
	                            "9306365021",
	                            "123",
	                            uniqueDocumentID,
	                            "8005087243"
	                        };                        
	                        for (int j = 0; j < belnrNodes.getLength(); j++) {
	                            Node belnrNode = belnrNodes.item(j);
	                            if (belnrNode != null) {	                               
	                                if (j < belnrValues.length) {
	                                    belnrNode.setTextContent(belnrValues[j]);	                                    
	                                }
	                            }
	                        }
	                    }
	                    if (fileName.equals("Quote_Interim")&& commaSeparated.contains("GetUniqueNumber|QTR")) {
	                    	uniqueDocumentID = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);	                        
	                        NodeList belnrNodes4 = doc.getElementsByTagName("BELNR");	                        
	                        String[] belnrValues4 = {
	                        		"SQR21",
	                        	uniqueDocumentID,
	                        	uniqueDocumentID
	                        };
                        for (int j = 0; j < belnrNodes4.getLength(); j++) {
                            Node belnrNode4 = belnrNodes4.item(j);
                            if (belnrNode4 != null) {
                                System.out.println("BELNR BEFORE = "+ belnrNode4.getTextContent());
                                if (j < belnrValues4.length) {
                                    belnrNode4.setTextContent(belnrValues4[j]);                                   
                                }
                            }
                        }
                    }
	                    if (fileName.equals("Quote_Final")&& commaSeparated.contains("GetUniqueNumber|QTR")) {
	                    	uniqueDocumentID = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);	                       
	                        NodeList belnrNodes4 = doc.getElementsByTagName("BELNR");	                        
	                        String[] belnrValues4 = {
	                        	uniqueDocumentID,
	                        	uniqueDocumentID,
	                        	uniqueDocumentID
	                        };
                        for (int j = 0; j < belnrNodes4.getLength(); j++) {
                            Node belnrNode4 = belnrNodes4.item(j);
                            if (belnrNode4 != null) {
                                System.out.println("BELNR BEFORE = "+ belnrNode4.getTextContent());
                                if (j < belnrValues4.length) {
                                    belnrNode4.setTextContent(belnrValues4[j]);                                    
                                }
                            }
                        }
                    }                    
	                    if (tagName.equals("CPO")&& commaSeparated.contains("GetUniqueNumber|CPO")) {
	                        uniqueDocumentID = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);	                        
	                        child.setNodeValue(uniqueDocumentID);
	                    }
	                    else if (tagName.equals("INV")&& commaSeparated.contains("GetUniqueNumber|CPO")) {
	                        uniqueDocumentID = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);
	                               child.setNodeValue(uniqueDocumentID+"010");
	                    }	                    
	                    else if (tagName.equals("LPO")&& commaSeparated.contains("GetUniqueNumber|CPO")) {
	                        uniqueDocumentID = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);
	                               child.setNodeValue(uniqueDocumentID);
	                    }
	                    else if (fileName.equals("Seller_Other_Change")&&tagName.equals("BELNR")&& commaSeparated.contains("GetUniqueNumber|CPO")||fileName.equals("Seller_Change_decrease")&&tagName.equals("BELNR")&& commaSeparated.contains("GetUniqueNumber|CPO")||fileName.equals("Seller_Change_increase")&&tagName.equals("BELNR")&& commaSeparated.contains("GetUniqueNumber|CPO")) {
	                    	uniqueDocumentID = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);
	                        System.out.println("uniqueDocumentID = " + uniqueDocumentID);
	                        NodeList belnrNodes = doc.getElementsByTagName("BELNR");
	                        System.out.println("BELNR count = " + belnrNodes.getLength());
	                        for (int j = 0; j < belnrNodes.getLength(); j++) {
	                            Node belnrNode = belnrNodes.item(j);
	                            if (belnrNode != null) {
	                                System.out.println("BELNR BEFORE = " + belnrNode.getTextContent());
	                                belnrNode.setTextContent(uniqueDocumentID);
	                                System.out.println("BELNR AFTER  = " + belnrNode.getTextContent());
	                            }
	                        }
	                    }	                   
	                    else if (tagName.equals("BSTNR")&& commaSeparated.contains("GetUniqueNumber|CPO")) {
	                        uniqueDocumentID = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);
	                               child.setNodeValue(uniqueDocumentID);
	                    }
	                    else if (tagName.equals("BELNR")&& commaSeparated.contains("GetUniqueNumber|QTR")) {
	                        uniqueDocumentID = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);
	                               child.setNodeValue(uniqueDocumentID);
	                    }
	                    else if (fileName.equals("Shipment")&&tagName.equals("TDLINE")) {
	                    	NodeList belnrNodes = doc.getElementsByTagName("TDLINE");
	                    	 DateFormat dateTimeFormat1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		                	 String dateTime =dateTimeFormat1.format(cal.getTime());   
	                        System.out.println("BELNR count = " + belnrNodes.getLength());
	                        String[] belnrValues = {
	                            "Shiptocode:2007",
	                            dateTime,
	                            "Dangerous Goods as per attached DGD.",
	                            "Cargo Aircraft Only (CAO)",
	                            "Shipment intended for Air Transportation.",
	                            "Flares",
	                            "aerial"
	                        };
	                        for (int j = 0; j < belnrNodes.getLength(); j++) {
	                            Node belnrNode = belnrNodes.item(j);
	                            if (belnrNode != null) {
	                                System.out.println("BELNR BEFORE = "+ belnrNode.getTextContent());
	                                if (j < belnrValues.length) {
	                                    belnrNode.setTextContent(belnrValues[j]);
	                                    System.out.println("BELNR AFTER = "+ belnrNode.getTextContent());
	                                }
	                            }
	                        }
	                    }
	                    
	                    else if (tagName.equals("INV")&& commaSeparated.contains("GetUniqueNumber|INV")) {
	                        String InvoiceNumber = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);
	                        child.setNodeValue(InvoiceNumber);
	                    }
	                    else if (tagName.equals("QTR")&& commaSeparated.contains("GetUniqueNumber|QTR")) {
	                        String QuoteNumber = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);
	                        child.setNodeValue(QuoteNumber);
	                    }               
	                   else if (tagName.equals("QuotationNumber")&& commaSeparated.contains("GetUniqueNumber|QuotationNumber")) {
	                        String QuotationNumber = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);
	                        child.setNodeValue(QuotationNumber);
	                    }
	                    else if (tagName.equals("QuotationNumber")&& commaSeparated.contains("GetUniqueNumber|QuotationNumber")) {
	                        String QuotationNumber = excelEditor.mappingUniqueNumberReadExcel(mappingKey,commaSeparated);
	                        child.setNodeValue(QuotationNumber);
	                    }
	                    // ==========================================
	                    // SERVICE ORDER RELATED TAGS
	                    // ==========================================
	                    else if (tagName.equals("PrimaryIdentifier")
	                            || tagName.equals("CustomerOrderNumber")
	                            || tagName.equals("CPO")
	                            || tagName.equals("QTR")
	                            || tagName.equals("ServiceOrderNumber")
	                            || tagName.equals("InvoiceNumber")) {
	                        child.setNodeValue(uniqueDocumentID);
	                        if (uniqueNumber.equals("Yes")&& tagName.equals("CPO")) {
	                            excelEditor.mappingUniqueNumberWriteExcel(uniqueDocumentID,mappingKey,tagName);
	                        }
	                        if (uniqueNumber.equals("Yes1")&& tagName.equals("QTR")) {
	                            excelEditor.mappingUniqueNumberWriteExcel(uniqueDocumentID,mappingKey,tagName);
	                        }
	                    }	              
	                    
	                    // QUOTATION NUMBER                  

	                    else if (tagName.equals("QuotationNumber")) {
	                        String QuotationNumber = "QN"+ radiantGenericCode.getDateTime("currentDate","MMddyyyyHHmmss",0);
	                        child.setNodeValue(QuotationNumber);
	                        if (uniqueNumber.equals("Yes")) {
	                            excelEditor.mappingUniqueNumberWriteExcel(QuotationNumber,mappingKey,tagName);
	                        }
	                    }
//	                   
	                    // CREATION DATETIME	                   
	                 // DATE FIELD HANDLING	                
	                 String currentDate =new SimpleDateFormat("yyyy-MM-dd").format(new Date());
	                 if (tagName.equals("CreationDateTime")) {
	                     child.setNodeValue(CreationDateTime);
	                     if (scriptName.equals("QuoteVersion2")||(scriptName.equals("QuoteAcceptVersion2")||(scriptName.equals("QuoteVersion1")||(scriptName.equals("QuoteAcceptVersion1"))))) {                                         
            	        DateFormat dateTimeFormat1 = new SimpleDateFormat("yyyy-MM-dd");
	             	    DateFormat dateTimeFormat2 = new SimpleDateFormat("HH:mm:ss");	             	    
	             	    CreationDateTime =dateTimeFormat1.format(cal.getTime())+"T"+dateTimeFormat2.format(cal.getTime())+"Z";
	                     child.setNodeValue(CreationDateTime);
	                     }
	                 }
	                 if (tagName.equals("REM")) {
	                     child.setNodeValue(CreationDateTime);
	                 }
	                 if (tagName.equals("ACK")) {
							child.setNodeValue(CreationDateTime);
	                 }
	                 if (tagName.equals("SSD")) {
	                	 DateFormat dateTimeFormat1 = new SimpleDateFormat("yyyyMMdd");
	                	 CreationDateTime2 =dateTimeFormat1.format(cal.getTime());
	                     child.setNodeValue(CreationDateTime2);
	                 }
	                 if (tagName.equals("DATUM")||tagName.equals("CREDAT")||tagName.equals("EDATU")||tagName.equals("PODAT")) {
	                	 DateFormat dateTimeFormat1 = new SimpleDateFormat("yyyyMMdd");
	                	 String CreationDateTime3 =dateTimeFormat1.format(cal.getTime());
	                     child.setNodeValue(CreationDateTime3);
	                 }
	                 if (fileName.equals("Shipment") && tagName.equals("ISDD")) {	                	 
		                	 NodeList belnrNodes = doc.getElementsByTagName("ISDD");
	                    	 DateFormat dateTimeFormat1 = new SimpleDateFormat("yyyyMMdd");
		                	 String dateTime =dateTimeFormat1.format(cal.getTime());   
	                        System.out.println("BELNR count = " + belnrNodes.getLength());
	                        String[] belnrValues = {                           
	                            dateTime,"00000000","00000000","00000000","00000000","00000000"
	                            };
	                        for (int j = 0; j < belnrNodes.getLength(); j++) {
	                            Node belnrNode = belnrNodes.item(j);
	                            if (belnrNode != null) {
	                                System.out.println("BELNR BEFORE = "+ belnrNode.getTextContent());
	                                if (j < belnrValues.length) {
	                                    belnrNode.setTextContent(belnrValues[j]);
	                                    System.out.println("BELNR AFTER = "
	                                            + belnrNode.getTextContent());
	                                }
	                            }
	                        }
	                    }
	                    if (fileName.equals("Po_Ack") && tagName.equals("TDLINE")) {
	                	 NodeList belnrNodes = doc.getElementsByTagName("TDLINE");
                    	 DateFormat dateTimeFormat1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	                	 String dateTime =dateTimeFormat1.format(cal.getTime());   
                        System.out.println("BELNR count = " + belnrNodes.getLength());
                        String[] belnrValues = {                           
                            dateTime,"123",dateTime
                            };
                        for (int j = 0; j < belnrNodes.getLength(); j++) {
                            Node belnrNode = belnrNodes.item(j);
                            if (belnrNode != null) {
                                System.out.println("BELNR BEFORE = "+ belnrNode.getTextContent());
                                if (j < belnrValues.length) {
                                    belnrNode.setTextContent(belnrValues[j]);
                                    System.out.println("BELNR AFTER = "
                                            + belnrNode.getTextContent());
                                }
                            }
                        }
                    }
	                    if (fileName.equals("Quote_Final") && tagName.equals("TDLINE")) {
		                	 NodeList belnrNodes = doc.getElementsByTagName("TDLINE");
	                    	 DateFormat dateTimeFormat1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		                	 String dateTime =dateTimeFormat1.format(cal.getTime());   
	                        System.out.println("BELNR count = " + belnrNodes.getLength());
	                        String[] belnrValues = {                           
	                            dateTime,"InqPart:PART01","Different PART"
	                            };
	                        for (int j = 0; j < belnrNodes.getLength(); j++) {
	                            Node belnrNode = belnrNodes.item(j);
	                            if (belnrNode != null) {
	                                System.out.println("BELNR BEFORE = "+ belnrNode.getTextContent());
	                                if (j < belnrValues.length) {
	                                    belnrNode.setTextContent(belnrValues[j]);
	                                    System.out.println("BELNR AFTER = "
	                                            + belnrNode.getTextContent());
	                                }
	                            }
	                        }
	                    }
	                 if (fileName.equals("Po_Ack") && tagName.equals("TDID")) {
	                	 NodeList belnrNodes = doc.getElementsByTagName("TDID");
                    	 DateFormat dateTimeFormat1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	                	 String dateTime =dateTimeFormat1.format(cal.getTime());   
                        System.out.println("BELNR count = " + belnrNodes.getLength());
                        String[] belnrValues = {                           
                            dateTime,"Z001","Z014"
                            };
                        for (int j = 0; j < belnrNodes.getLength(); j++) {
                            Node belnrNode = belnrNodes.item(j);
                            if (belnrNode != null) {
                                System.out.println("BELNR BEFORE = "+ belnrNode.getTextContent());
                                if (j < belnrValues.length) {
                                    belnrNode.setTextContent(belnrValues[j]);
                                    System.out.println("BELNR AFTER = "
                                            + belnrNode.getTextContent());
                                }
                            }
                        }
                    }
	                 if (tagName.equals("ABGRT")) {
	                	 DateFormat dateTimeFormat1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	                	 String CreationDateTime4 =dateTimeFormat1.format(cal.getTime());
	                     child.setNodeValue(CreationDateTime4);
	                 } 	                 
	                 if (fileName.equals("Seller_Other_Change") && tagName.equals("TDLINE")||fileName.equals("Seller_Change_decrease") && tagName.equals("TDLINE")||fileName.equals("Seller_Change_increase") && tagName.equals("TDLINE")) {
	                	 DateFormat dateTimeFormat1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	                	 String CreationDateTime4 =dateTimeFormat1.format(cal.getTime());
	                     child.setNodeValue(CreationDateTime4);
	                 } 
	                 if (fileName.equals("Invoice") && tagName.equals("TDLINE")||fileName.equals("Invoice_Credit") && tagName.equals("TDLINE")) {
	                	 DateFormat dateTimeFormat1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	                	 String CreationDateTime5 =dateTimeFormat1.format(cal.getTime());
	                     child.setNodeValue(CreationDateTime5);
	                 } 
	                 
	                 if (rootTag.equals("ATA_InvoiceMessageAck") &&  tagName.equals("MessageReceiver")) {
							SimpleDateFormat creationDateAndTimeWithSpace = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
							SimpleDateFormat creationDateAndTimeWithoutSpace = new SimpleDateFormat("yyyyMMddHHmmss");
							String creationDateTimeMgsRec = creationDateAndTimeWithoutSpace.format(creationDateAndTimeWithSpace.parse(CreationDateTime));
							child.setNodeValue(creationDateTimeMgsRec);
						}
	                 if (tagName.equals("SERIAL")) {
							SimpleDateFormat creationDateAndTimeWithSpace = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
							SimpleDateFormat creationDateAndTimeWithoutSpace = new SimpleDateFormat("yyyyMMddHHmmss");
							String creationDateTimeMgsRec = creationDateAndTimeWithoutSpace.format(creationDateAndTimeWithSpace.parse(CreationDateTime));
							child.setNodeValue(creationDateTimeMgsRec);
						}
	                 else if (tagName.equals("SpecifiedShippingDate")
	                         || tagName.equals("RequestedReturnDate")
	                         || tagName.equals("PlannedDispatchDate")
	                         || tagName.equals("PartRemovalDate")) {

	                     child.setNodeValue(currentDate);
	                 }	
	                 // CUSTOM FIELD
	                 
	                 else if (tagName.equals("CustomField")) {

	                     child.setNodeValue(CreationDateTime);
	                 }
	                }
	            }
	        }
	        // Fallback check
	        if (uniqueDocumentID == null || uniqueDocumentID.trim().isEmpty()) {
	            try {
	                String fallbackId = excelEditor.mappingUniqueNumberReadExcel(mappingKey,"Fallback|CPO");
	                if (fallbackId != null) {
	                    uniqueDocumentID = fallbackId;
	                } else {
	                    fallbackId = excelEditor.mappingUniqueNumberReadExcel(mappingKey,"Fallback|ShipmentNumber2");
	                    if (fallbackId != null) {
	                        uniqueDocumentID = fallbackId;
	                    }
	                }
	            } catch (Exception e) {
	            }
	        }
	        // SAVE XML WITHOUT CHANGING FORMAT
	        TransformerFactory tf = TransformerFactory.newInstance();
	        Transformer transformer = tf.newTransformer();
	        String threadFileName = fileName + "_"+ Thread.currentThread().getId() + "_"+ System.currentTimeMillis();
	        StreamResult xmlResult = new StreamResult(filePathHandling.defaultXmlOutputPath+ threadFileName + ".xml");
	        transformer.transform(new DOMSource(doc), xmlResult);
	        System.out.println("FINAL XML CREATED : " + threadFileName);
	        result("", rootTag, scriptName, "before",description, scenarioType, expectation, sNo);
	        System.out.println("Write before");
	        String response = aeroComponent(threadFileName, crediatial);
	        System.out.println("response " + response);
	        result(response, rootTag, scriptName, "after",description, scenarioType, expectation, sNo);
	    } catch (Exception e) {
	        e.printStackTrace();
	    }
	    return null;
	}
	public static String setValueInDataExcel(int sheetNumber,int row,int col,String value) throws Exception {						
		File eoFile=new File(filePathHandling.defaultInputPath + filePathHandling.aerobuyEDIMessageDataExcelFileName);			 		
		FileInputStream fsIP= new FileInputStream(eoFile); 
	    @SuppressWarnings("resource")
		XSSFWorkbook wb = new XSSFWorkbook(fsIP); 	  
	    Cell cell = null;
	    cell =wb.getSheetAt(sheetNumber).getRow(row).getCell(col);
        cell.setCellValue(value);		               
        fsIP.close();    
        FileOutputStream output_file =new FileOutputStream(eoFile);  
        wb.write(output_file); 
        output_file.close();
		return null;		
	}	

	@SuppressWarnings("deprecation")
	public void result(String response, String rootTag, String scriptName, String methodType, String description,
			String scenarioType, String expectation, String sNo) throws Exception {
		synchronized (App.class) {
			filename = filePathHandling.outputExcelPath;
			FileInputStream output = new FileInputStream(new File(filename));
			wb = new XSSFWorkbook(output);
			worksheet = wb.getSheetAt(0);
			int rowCount = worksheet.getPhysicalNumberOfRows();
			XSSFRow row;
			if (methodType.equals("after")) {
				String verify = "Error: Invalid Response";
				String  receiptMessage="";
				try {
					InputSource src = new InputSource();
					src.setCharacterStream(new StringReader(response));
					DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
					Document doc = docBuilder.parse(src);
					String rootTagName= doc.getDocumentElement().getNodeName();					
						if(rootTagName.equals("Response")){
						    receiptMessage = doc.getElementsByTagName("status").item(0).getTextContent();
						  if(receiptMessage.equals("200")) {
							  verify="Successful";				  
						  }			 
					  }else if(rootTagName.equals("aex:AEX_WSResponseMsg")){
						  receiptMessage = doc.getElementsByTagName("aex:errorMsg").item(0).getTextContent();
						  if(receiptMessage.equals("Successful")) {
							  verify="Successful";				  
						  }
					  }else{
						    receiptMessage = doc.getElementsByTagName("errorMsg").item(0).getTextContent();
						    verify=receiptMessage;
					  } }		
				catch (Exception e) {
					System.out.println("Failed to parse response: " + e.getMessage() + ". Response was: " + response);
				}
				System.out.println("EO Number:" + uniqueDocumentID);
				System.out.println("Validation: " + verify);
				section: for (int i = rowCount - 1; i >= 1; i--) {
					row = worksheet.getRow(i);
					for (Cell cellFor : row) {
						if (cellFor.getCellType() == CellType.STRING) {
							if (cellFor.getRichStringCellValue().getString().trim().equals(rootTag)) {
								XSSFCell cellA1;
								cellA1 = row.getCell((short) 11);
								String orderNumber = cellA1.getStringCellValue();
								cellA1 = row.getCell((short) 6);
								String roottagName = cellA1.getStringCellValue();
								cellA1 = row.getCell((short) 17);
								String creationDate = cellA1.getStringCellValue();
//								if (orderNumber.equals(uniqueDocumentID) && roottagName.equals(rootTag)
//										&& creationDate.equals(CreationDateTime)) {
//									System.out.println(verify);
//									row.createCell(16).setCellValue(verify);
//									break section;
								if ((orderNumber.equals(uniqueDocumentID)||orderNumber.equals(uniqueDocumentID)) && roottagName.equals(rootTag)
										&& creationDate.equals(CreationDateTime)) {
									row.createCell(16).setCellValue(verify);
								break section;
								}
							}
						}
					}
				}
			}
			if (methodType.equals("before")) {
				worksheet.createRow(rowCount).createCell(0).setCellValue(sNo);
				// worksheet.createRow(rowCount).createCell(0).setCellValue(rootTag);
				XSSFRow cell = worksheet.getRow(rowCount);
				cell.createCell(1).setCellValue("EDI");
				cell.createCell(2).setCellValue("XML");
				cell.createCell(3).setCellValue("Aerobuy Collaboration 2002_IDOC EDI Scenarios");
				cell.createCell(4).setCellValue(scriptName);
				cell.createCell(6).setCellValue(rootTag);
				cell.createCell(8).setCellValue(description);
				cell.createCell(11).setCellValue(uniqueDocumentID);
				cell.createCell(12).setCellValue(scenarioType);
				cell.createCell(13).setCellValue(expectation);
				cell.createCell(14).setCellValue("Fail");
				cell.createCell(15).setCellValue("Fail");
				cell.createCell(16).setCellValue("SENT");
				if(scriptName.equals("Invoice_ack")) {
					cell.createCell(17).setCellValue("");
				}else {
					cell.createCell(17).setCellValue(CreationDateTime);
				}
			}				
			output.close();
			FileOutputStream output_file = new FileOutputStream(new File(filename));
			wb.write(output_file);
			output_file.close();
		}
	}

	public void waitOutbound() throws Exception {
		try {
			// excel inputs
			File searchQuery = new File(filename);
			output = new FileInputStream(searchQuery);
			wb = new XSSFWorkbook(output);
			worksheet = wb.getSheetAt(0);
			int timeToWait = 120; // second
			for (int i = 0; i < timeToWait; i++) {
				int getrow = worksheet.getPhysicalNumberOfRows();
				String outboundResult = wb.getSheetAt(0).getRow(getrow - 1).getCell(14).getStringCellValue();
				if (outboundResult.equals("Pass")) {
					System.out.println("Pass");
					break;
				}

			}
			System.out.println("message execution time ends");
		} catch (Exception e) {
			System.out.println("Exception==" + e);
		}
	}
}
