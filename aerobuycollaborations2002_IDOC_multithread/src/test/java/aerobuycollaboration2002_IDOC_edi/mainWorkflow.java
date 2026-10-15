package aerobuycollaboration2002_IDOC_edi;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import aerobuycollaboration2002_IDOC_edi.App;
import aerobuycollaboration2002_IDOC_edi.excelEditor;
import aerobuycollaboration2002_IDOC_edi.filePathHandling;

public class mainWorkflow {
    public static void workflowProcess(int threadCount) throws Exception {
        Map<String, List<scenarioData>> scenarioMap = readScenarioSheet(
                filePathHandling.defaultInputPath + filePathHandling.aerobuyEDIScenarioExcelFileName);
        List<Map<String, String>> workflowData = readWorkflowSheet(
                filePathHandling.defaultInputPath + filePathHandling.aerobuyEDIScenarioExcelFileName);

        // 1. Thread-safe Set for processed keys to prevent
        // ConcurrentModificationException
        Set<String> processedKeys = java.util.concurrent.ConcurrentHashMap.newKeySet();

        // 2. We use ONE executor for the Scenarios (the columns).
        // We CANNOT multithread the Workflow rows because they represent sequential
        // Steps in time!
        ExecutorService executor = Executors.newFixedThreadPool(threadCount);

        for (Map<String, String> row : workflowData) {
            System.out.println("Processing New Workflow Row (Sequential Step)");
            List<Future<?>> messageFutures = new ArrayList<>();
            int mappingUniqueNumberWriteExcel = 0;

            for (Map.Entry<String, String> entry : row.entrySet()) {
                String scenarioId = entry.getKey();
                String messageId = entry.getValue();
                // Skip empty cells
                if (messageId == null || messageId.trim().isEmpty()) {
                    continue;
                }
                // Create key
                String key = scenarioId + "|" + messageId;
                System.out.println("Workflow Entry → " + key);
                // Fetch ALL scenario data (List)
                List<scenarioData> dataList = scenarioMap.get(key);
                if (dataList == null || dataList.isEmpty()) {
                    System.out.println("No scenario found for key: " + key);
                    continue;
                }
                // Get XML File Name
                String xmlFileName = xmlNameReader.getXmlName(filePathHandling.defaultInputPath
                        + filePathHandling.aerobuyEDIScenarioExcelFileName, messageId);
                System.out.println("xmlFileName: " + xmlFileName);

                // xml editing and posting - these scenarios CAN run concurrently
                for (scenarioData data : dataList) {
                    System.out.println("Executing Mapping Key: " + key);
                    Map<String, String> rowData = data.getRowData();
                    String sNo = rowData.get("S.No");
                    String uniqueKey = scenarioId + "|" + messageId + "|" + sNo;
                    if (!processedKeys.add(uniqueKey)) {
                        System.out.println("Skipping duplicate: " + uniqueKey);
                        continue;
                    }
                    String messageName = rowData.get("Message Name");
                    String messageType = rowData.get("Message Type");
                    String elementNameRootTag = rowData.get("Element Name/Root Tag");
                    String uniqueNumber = rowData.get("Unique Number");
                    int sheetNumber = Integer.parseInt(rowData.get("Sheet Number"));
                    String description = rowData.get("Description");
                    String scenarioType = rowData.get("Scenario Type");
                    String expectation = rowData.get("Expectation");
                    String username = "aextn";
                    String password = "aextn";
                    String userPassword = username + ":" + password;
                    String base64Encoder = Base64.getEncoder().encodeToString(userPassword.getBytes());

                    // Effectively final variable for lambda
                    int currentMappingUniqueNumber = mappingUniqueNumberWriteExcel;

                    // Submit this scenario step to be processed concurrently
                    messageFutures.add(executor.submit(() -> {
                        try {
                            App mainClass = new App();
                            mainClass.xmlMapping(xmlFileName, sheetNumber, base64Encoder, elementNameRootTag,
                                    messageName, description, scenarioType, expectation, messageType,
                                    uniqueNumber, key, currentMappingUniqueNumber, sNo);
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }));
                }
            }

            // Wait for ALL scenarios in THIS SPECIFIC STEP to finish
            for (Future<?> f : messageFutures) {
                try {
                    f.get();
                } catch (Exception e) {
                }
            }

            // Write all Excel files safely
            synchronized (excelEditor.class) {
                System.out.println("mappingUniqueNumberWriteExcelClose");
                excelEditor.mappingUniqueNumberWriteExcelClose();
            }

            // CRITICAL: Wait 2 minutes before progressing to the Next Step (Row)
            // This ensures external systems process previous files (like Creation) fully
            System.out.println("Waiting for external system processing... 120s");
         //   Thread.sleep(120000);
        }

        // Shut down executor to free memory
        executor.shutdown();
        executor.awaitTermination(1, TimeUnit.HOURS);

        System.out.println("sorting started");
        Thread.sleep(5000);
        excelEditor.sortExcelSheetByColumn(filePathHandling.outputExcelPath, 0);
     //    Thread.sleep(600000);
        System.out.println("sorting ended");
    }

    public static List<Map<String, String>> readWorkflowSheet(String filePath) throws Exception {
        List<Map<String, String>> workflowRows = new ArrayList<>();
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);
        // 3rd sheet → index 2 (because index starts from 0)
        Sheet sheet = workbook.getSheetAt(3);
        // Read header row (row 0)
        Row headerRow = sheet.getRow(0);
        int lastRowNum = sheet.getLastRowNum();
        int lastCellNum = headerRow.getLastCellNum();
        // Loop from row 1 (skip header)
        for (int i = 1; i <= lastRowNum; i++) {
            Row currentRow = sheet.getRow(i);
            if (currentRow == null)
                continue;
            Map<String, String> rowData = new LinkedHashMap<>();
            // Loop all columns
            for (int j = 1; j < lastCellNum; j++) {
                // j=1 because we skip S.No column (index 0)
                Cell headerCell = headerRow.getCell(j);
                Cell dataCell = currentRow.getCell(j);
                if (headerCell == null)
                    continue;
                String columnName = headerCell.getStringCellValue();
                String cellValue = "";
                if (dataCell != null)
                    cellValue = dataCell.toString().trim();
                rowData.put(columnName, cellValue);
            }
            workflowRows.add(rowData);
        }
        workbook.close();
        fis.close();
        return workflowRows;
    }

    public static Map<String, List<scenarioData>> readScenarioSheet(String filePath) throws Exception {

        Map<String, List<scenarioData>> scenarioMap = new HashMap<>();

        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);

        // Scenario sheet (example: sheet index 2)
        Sheet sheet = workbook.getSheetAt(2);

        Row headerRow = sheet.getRow(0);
        int lastRow = sheet.getLastRowNum();
        int lastCell = headerRow.getLastCellNum();

        for (int i = 1; i <= lastRow; i++) {

            Row row = sheet.getRow(i);
            if (row == null)
                continue;

            Map<String, String> rowData = new HashMap<>();

            for (int j = 0; j < lastCell; j++) {

                Cell headerCell = headerRow.getCell(j);
                Cell dataCell = row.getCell(j);

                String columnName = headerCell.getStringCellValue();
                String cellValue = "";

                if (dataCell != null)
                    cellValue = dataCell.toString().trim();

                rowData.put(columnName, cellValue);
            }

            // Read Scenario ID and Message ID
            String scenarioId = rowData.get("Scenario ID");
            String messageId = rowData.get("Message ID");

            // Skip invalid rows
            if (scenarioId == null || messageId == null ||
                    scenarioId.trim().isEmpty() || messageId.trim().isEmpty()) {
                continue;
            }

            // Create key
            String key = scenarioId + "|" + messageId;

            scenarioData data = new scenarioData();
            data.setRowData(rowData);

            // ✅ IMPORTANT FIX: store multiple entries per key
            scenarioMap
                    .computeIfAbsent(key, k -> new ArrayList<>())
                    .add(data);
        }

        workbook.close();
        fis.close();

        return scenarioMap;
    }
}
