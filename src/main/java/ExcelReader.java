import java.io.File;
import java.util.*;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Iterator;

public class ExcelReader {

    // Default constructor for this object
    public ExcelReader() {
    }

    // This function parses the given file and returns
    // an ordered list of StartEnd objects that
    // represent the intervals within the first column
    // of the Excel Sheet
    public List<StartEnd> parseExcel(String fileName) {
        List<Integer> listOfCodes = new ArrayList<>();
        try {
            String currentDirectory = System.getProperty("user.dir");
            System.out.println("The current working directory is " + currentDirectory);
            FileInputStream file = new FileInputStream(new File(fileName));

            // Create Workbook instance holding reference to .xlsx file
            // Workbook workbook = WorkbookFactory.create(file);
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get first/desired sheet from the workbook
            // In this case, I'll only consider one sheet
            // Otherwise we can loop through the sheets as well
            XSSFSheet sheet = workbook.getSheetAt(0);

            // Iterate through each row one by one
            // Once again, assuming there's only one row here.
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                        // Add the cell into our list
                        listOfCodes.add((int) cell.getNumericCellValue());
                    } else {
                        // shouldn't happen, try to just skip past it
                        continue;
                    }
                }
            }
            file.close();
        }
        catch (Exception e) {
            e.printStackTrace();
        }

        // Once we are done adding all the columns, sort our list and begin
        // creating our list of StartEnd objects. Also note that
        // it uses quicksort on integers so it's as fast as possible.
        Collections.sort(listOfCodes);
        List<StartEnd> res = new ArrayList<>();

        for (int i = 0; i < listOfCodes.size(); i++) {
            StartEnd temp = new StartEnd();
            temp.setStart(listOfCodes.get(i));

            for (int j = i + 1; j < listOfCodes.size(); j++) {
                if (j == listOfCodes.size() - 1) {
                    // at the end, check which case we are
                    // case 1: continues the current StartEnd
                    if (listOfCodes.get(j) == listOfCodes.get(i) + 1) {
                        temp.setEnd(listOfCodes.get(j));
                        res.add(temp);
                        i = j;
                        break;
                    } else {
                        // case 2: it doesn't continue; it's
                        // on it's own
                        temp.setEnd(listOfCodes.get(i));
                        res.add(temp);
                        int last = listOfCodes.get(j);
                        res.add(new StartEnd(last, last));
                        break;
                    }
                }
                if (listOfCodes.get(j) != listOfCodes.get(j - 1) + 1) {
                    temp.setEnd(listOfCodes.get(i));
                    res.add(temp);
                    break;
                } else {
                    i++;
                }
            }
        }
        return res;
    }

    public void print(List<StartEnd> res) {
        printList(res);
    }

    private void printList(List<StartEnd> res) {
        for (int i = 0; i < res.size(); i++) {
            StartEnd print = res.get(i);
            System.out.println("(" + print.getStart() + "," + print.getEnd() + ")");
        }
    }
}
