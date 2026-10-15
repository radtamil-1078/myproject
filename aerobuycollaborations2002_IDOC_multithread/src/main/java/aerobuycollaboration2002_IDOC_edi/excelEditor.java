package aerobuycollaboration2002_IDOC_edi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class excelEditor {
	public static int shipmentRowIndex = 1;
	public static int orderNumberIndex = 1;
	public static String dateFolder;
//	public static FileInputStream mappingDocument;
//	public static XSSFWorkbook mappingWorkbook;
//	public static XSSFSheet mappingSheet;
	public static int mappingUniqueNumberWriteExcelcount=0;
	public static String currentDate;	
//	 private static String mappingDocumentPath;
	    public static XSSFWorkbook mappingWorkbook;
	    public static XSSFSheet mappingSheet;
	    private static String mappingDocumentPath;

	    public static synchronized void mappingUniqueNumberWriteExcelOpen(String filePath) throws Exception {
	        mappingDocumentPath = filePath;
	        File file = new File(mappingDocumentPath);
	        if (!file.exists() || file.length() == 0) {
	            System.out.println("Excel not created. Creating new file: " + mappingDocumentPath);
	            mappingWorkbook = new XSSFWorkbook();
	            mappingSheet = mappingWorkbook.createSheet("Mapping");
	            // Create Header
	            XSSFRow header = mappingSheet.createRow(0);
	            header.createCell(0).setCellValue("S.No");
	            header.createCell(1).setCellValue("UniqueNumber");

	            // Save immediately after creation
	            FileOutputStream fos = new FileOutputStream(file);
	            mappingWorkbook.write(fos);
	            fos.close();

	        } else {
	            FileInputStream fis = new FileInputStream(file);
	            mappingWorkbook = new XSSFWorkbook(fis);
	            fis.close();  // VERY IMPORTANT
	            mappingSheet = mappingWorkbook.getSheetAt(0);
	        }
	    }

	    public static synchronized void mappingUniqueNumberWriteExcelClose() throws Exception {
	        if (mappingWorkbook != null) {
	            FileOutputStream fos = new FileOutputStream(mappingDocumentPath);
	            mappingWorkbook.write(fos);
	            fos.close();
	            mappingWorkbook.close();
	            mappingWorkbook = null;
	            mappingSheet = null;
	            mappingUniqueNumberWriteExcelcount=0;
	            System.out.println("Excel closed successfully.");
	        }
	    }

	    public static synchronized void mappingUniqueNumberWriteExcel(String fieldData,String mappingKey,String tagName) throws Exception {	    	
	    	if(mappingUniqueNumberWriteExcelcount==0) {
	    		mappingUniqueNumberWriteExcelOpen(filePathHandling.mappingExcelPath);
	    		mappingUniqueNumberWriteExcelcount=1;
			}	
	        if (mappingWorkbook == null) {
	            throw new IllegalStateException("Workbook not opened. Call open() first.");
	        }
	        int lastRow = mappingSheet.getLastRowNum();
	        int newRowIndex = lastRow + 1;
	        XSSFRow row = mappingSheet.createRow(newRowIndex);
	        row.createCell(0).setCellValue(newRowIndex);
	        String scenarioCode = mappingKey.split("\\|")[0];
	        row.createCell(1).setCellValue(scenarioCode +"|" +tagName+"|" + fieldData);
	    }
	    public static synchronized String mappingUniqueNumberReadExcel(String mappingKey,String commaSeparated) throws Exception {			
	    	if (mappingUniqueNumberWriteExcelcount == 0) {
	    	    mappingUniqueNumberWriteExcelOpen( filePathHandling.mappingExcelPath);
	    	    mappingUniqueNumberWriteExcelcount = 1;
	    	}
	    	// Extract SC01 from mappingKey
	    	String scenarioCode = mappingKey.split("\\|")[0];
	    	// Extract OrderNumber from commaSeparated
	    	String fieldName = commaSeparated.split("\\|")[1];
	    	// Merge both
	    	String finalKey = scenarioCode + "|" + fieldName;
	    	for (int i = 1; i <= mappingSheet.getLastRowNum(); i++) { 
	    		// Skip header
	    	    Row row = mappingSheet.getRow(i);
	    	    if (row == null)
	    	        continue;
	    	    Cell cell = row.getCell(1); // UniqueNumber column
	    	    if (cell == null)
	    	        continue;
	    	    String cellValue = cell.getStringCellValue().trim();
	    	    // Match using finalKey
	    	    if (cellValue.startsWith(finalKey + "|")) {
	    	        String[] parts = cellValue.split("\\|");
	    	        if (parts.length == 3) {
	    	            return parts[2]; // VALxxxxxxxx
	    	        }
	    	    }
	    	}
	    	return null;
	    }
	    public static void sortExcelSheetByColumn(String filePath, int columnIndex) throws Exception {
	        FileInputStream fis = new FileInputStream(filePath);
	        Workbook workbook = new XSSFWorkbook(fis);
	        Sheet sheet = workbook.getSheetAt(0);
	        List<List<String>> data = new ArrayList<>();
	        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
	            Row row = sheet.getRow(i);
	            List<String> rowData = new ArrayList<>();
	            for (int j = 0; j < row.getLastCellNum(); j++) {
	                Cell cell = row.getCell(j);
	                rowData.add(cell == null ? "" : cell.toString());
	            }
	            data.add(rowData);
	        }
	        // Sort using S.No column
	        data.sort(Comparator.comparingInt(r -> Integer.parseInt(r.get(columnIndex))));
	        // Remove old rows
	        for (int i = sheet.getLastRowNum(); i > 0; i--) {
	            sheet.removeRow(sheet.getRow(i));
	        }
	        // Write sorted rows
	        int rowNum = 1;
	        for (List<String> rowData : data) {
	            Row row = sheet.createRow(rowNum++);
	            for (int j = 0; j < rowData.size(); j++) {
	                row.createCell(j).setCellValue(rowData.get(j));
	            }
	        }
	        fis.close();
	        FileOutputStream fos = new FileOutputStream(filePath);
	        workbook.write(fos);
	        fos.close();
	        workbook.close();
	    }
	    
	public static void verifyingOutputExcel() {		
		String mainOutputFile = filePathHandling.outputExcelPath;
        String validationRequiredOrderNumberDetailsExcel = filePathHandling.validationOrderDetailExcelPath;
        DataFormatter formatter = new DataFormatter();
        try (FileInputStream outputExcelFis = new FileInputStream(mainOutputFile);
             FileInputStream validationOrderNumberDetailsExcelFis = new FileInputStream(validationRequiredOrderNumberDetailsExcel);
             XSSFWorkbook outputExcelWb = new XSSFWorkbook(outputExcelFis);
             XSSFWorkbook validationOrderNumberDetailsExcelWb = new XSSFWorkbook(validationOrderNumberDetailsExcelFis)) {

            XSSFSheet outputExcelSheet = outputExcelWb.getSheetAt(0);
            XSSFSheet validationSheet = validationOrderNumberDetailsExcelWb.getSheetAt(0);

            // STEP 1: Preload validation rows into a Map for O(1) lookup
            Map<String, String> validationMap = new HashMap<>();
            System.out.println("validationSheet.getLastRowNum() " +validationSheet.getLastRowNum());
            for (int j = 1; j <= validationSheet.getLastRowNum(); j++) {
                Row row = validationSheet.getRow(j);
                if (row == null) continue;

                String rootTag = formatter.formatCellValue(row.getCell(1));
                String orderNumber = formatter.formatCellValue(row.getCell(2));
                String postedDateTime = formatter.formatCellValue(row.getCell(3));
                String receivedDateTime = formatter.formatCellValue(row.getCell(4));

                // Create a composite key for fast match
                String key = rootTag + "|" + orderNumber + "|" + postedDateTime;
                validationMap.put(key, receivedDateTime);
            }
            System.out.println("Total rows loaded in map: " + validationMap.size());

            // STEP 2: Process output rows and update where match found
            for (int i = 1; i <= outputExcelSheet.getLastRowNum(); i++) {
                Row outputRow = outputExcelSheet.getRow(i);
                if (outputRow == null) continue;

                String orderNumber = formatter.formatCellValue(outputRow.getCell(11));
                String postedDateTime = formatter.formatCellValue(outputRow.getCell(17));
                String rootTag = formatter.formatCellValue(outputRow.getCell(6));
                String expectation = formatter.formatCellValue(outputRow.getCell(13));
                String actualResult = formatter.formatCellValue(outputRow.getCell(14));

                if ("Fail".equalsIgnoreCase(actualResult)) {
                    String key = rootTag + "|" + orderNumber + "|" + postedDateTime;
                    String receivedDateTime = validationMap.get(key);

                    if (receivedDateTime != null) {
                        outputRow.createCell(14).setCellValue("Pass");
                        outputRow.createCell(18).setCellValue(receivedDateTime);
                        actualResult = "Pass";
                    }
                }

                if (expectation.equalsIgnoreCase(actualResult)) {
                    outputRow.createCell(15).setCellValue("Pass");
                }
            }

            // STEP 3: Write to output only once
            try (FileOutputStream fos = new FileOutputStream(mainOutputFile)) {
                outputExcelWb.write(fos);
            }

            System.out.println("✅ Excel processing completed successfully!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }        
       
}
