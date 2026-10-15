package aerobuycollaboration2002_IDOC_edi;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class xmlNameReader {

    public static String getXmlName(String filePath, String messageId) throws Exception {

        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(1);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row.getCell(1).getStringCellValue().equals(messageId)) {
                String xmlName = row.getCell(4).getStringCellValue();
                String xmlRootTag = row.getCell(2).getStringCellValue();                
                workbook.close();
                return xmlName;
            }
        }

        workbook.close();
        return null;
    }
}
