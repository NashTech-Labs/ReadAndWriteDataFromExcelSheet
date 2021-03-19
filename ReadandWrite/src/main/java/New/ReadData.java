package New;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;


public class ReadData
{
    @Test
    public void Read() throws EncryptedDocumentException, IOException {
        //get the excel sheet file location
        FileInputStream fis = new FileInputStream("src//test//java//Read.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        //We give the sheet index for fetching the data
        //  XSSFSheet sheet = workbook.getSheetAt(0);
        // We give the sheet name for fetching the data
        XSSFSheet sheet = workbook.getSheet(" Student Data ");
        //get the total row count in the excel sheet
        Iterator<Row> iterator = sheet.rowIterator();
        int cellIndex = 0;
        String description = null;
        while (iterator.hasNext()) {
            Row row = iterator.next();
            for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
                Cell cell = row.getCell(i);
                //print the cell value
                System.out.println(i + " " +cell );

            }
        }
    }
}