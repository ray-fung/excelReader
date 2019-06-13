import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

import static org.junit.Assert.*;

public class ExcelReaderTest {

    // Private variables for testing
    private ExcelReader test;

    @Before
    public void setUp() throws Exception {
        // Create a custom workbook
        try {
            FileOutputStream file = new FileOutputStream(new File("temp1.xlsx"));

            // Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook();
            // Create a sheet to test
            Sheet sheet = workbook.createSheet("Test");

            // All we really need to test is populating the first row.
            Object[] data = {1,2,3,4,5,7,8,9,10,"HELLO",11,13,13,20,54,32,31,30,19,21,18,40,70,71,72,99,105};
            // Note: I've included a string in this so we can test if the
            // reader successfully ignores it
            int rowCount = 0;

            for (Object dat : data) {
                Row row = sheet.createRow(++rowCount);
                // Just write all the values of data into the field
                Cell cell = row.createCell(0);
                if (dat instanceof String) {
                    cell.setCellValue((String) dat);
                } else if (dat instanceof Integer) {
                    cell.setCellValue((Integer) dat);
                }
            }
            workbook.write(file);
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Just going to try to create a blank worksheet and make sure
        // excelReader reads nothing
        try {
            FileOutputStream file = new FileOutputStream(new File("temp2.xlsx"));

            // Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook();
            // Create a sheet to test
            Sheet sheet = workbook.createSheet("Test");

            // All we really need to test is populating the first row.
            Object[] data = {};
            // Note: I've included a string in this so we can test if the
            // reader successfully ignores it
            int rowCount = 0;

            for (Object dat : data) {
                Row row = sheet.createRow(++rowCount);
                // Just write all the values of data into the field
                Cell cell = row.createCell(0);
                if (dat instanceof String) {
                    cell.setCellValue((String) dat);
                } else if (dat instanceof Integer) {
                    cell.setCellValue((Integer) dat);
                }
            }
            workbook.write(file);
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        test = new ExcelReader();
    }

    @After
    public void tearDown() throws Exception {
        // Delete the two Excel files that were created
        File temp1 = new File("temp1.xlsx");
        if (temp1.delete()){
            System.out.println("temp1.xlsx file deleted from Project root directory");
        } else
            System.out.println("File temp1.xlsx doesn't exist in the project root directory");

        File temp2 = new File("temp2.xlsx");
        if (temp2.delete()){
            System.out.println("temp2.xlsx File deleted from Project root directory");
        } else
            System.out.println("File temp2.xlsx doesn't exist in the project root directory");
    }

    @Test
    public void test1() {

        System.out.println("Testing populated file: ");

        // Test the base file we have created here
        List<StartEnd> temp = test.parseExcel("temp.xlsx");

        // Print the results
        test.print(temp);
        // There should be 10 different StartEnds
        assertEquals(11, temp.size());
    }

    @Test
    public void test2() {

        System.out.println("Testing empty file: ");

        // Test a blank excel file
        List<StartEnd> temp = test.parseExcel("temp2.xlsx");
        // Print the results (should be empty actually)
        test.print(temp);
        // There should be 0 StartEnd objects
        assertEquals(0, temp.size());
    }

}