package excelReaderTest;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {
    private final String excelPath;
    private XSSFSheet sheet;
    private XSSFWorkbook book;
    private String sheetName;

    public ExcelReader(String excelPath) {
        this.excelPath = excelPath;
        File file = new File(excelPath);
        try {
            FileInputStream fileInputStream = new FileInputStream(file);
            book = new XSSFWorkbook(fileInputStream);
            sheet = book.getSheet("Лист1");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public ExcelReader(String excelPath, String sheetName) {
        this.excelPath = excelPath;
        this.sheetName = sheetName;
        File file = new File(excelPath);
        try {
            FileInputStream fileInputStream = new FileInputStream(file);
            book = new XSSFWorkbook(fileInputStream);
            sheet = book.getSheet(sheetName);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private String cellToString(XSSFCell cell) {
        Object result;
        CellType type = cell.getCellType();
        switch (type) {
            case NUMERIC:
                result = cell.getNumericCellValue();
                break;
            case STRING:
                result = cell.getStringCellValue();
                break;
            case FORMULA:
                result = cell.getCellFormula();
                break;
            case BLANK:
                result = "";
                break;
            default:
                throw new IllegalStateException("ReadFormat Exception");
        }
        return result.toString();
    }

    private int xlsxColumnCount() {
        return sheet.getRow(1).getLastCellNum();
    }

    private int xlsxRowCount() {
        return sheet.getLastRowNum()+1;
    }

    public String[][] getDataSheetForTDD() throws IOException {
        File file = new File(excelPath);
        FileInputStream fileInputStream = new FileInputStream(file);
        book = new XSSFWorkbook(fileInputStream);
        sheet = book.getSheet("Лист1");
        int rowNumber = xlsxRowCount();
        int columnNumber = xlsxColumnCount();
//        System.out.println("columnNumber = " + columnNumber);
//        System.out.println("rowNumber = " + rowNumber);
        String[][] data = new String[rowNumber][columnNumber];
        for (int i = 1; i < rowNumber; i++) {
            for (int j = 0; j < columnNumber; j++) {
                XSSFRow row = sheet.getRow(i);
                XSSFCell cell = row.getCell(j);
                String value = cellToString(cell);
                data[i][j] = value;
                if (value == null) {
                    System.out.println("Empty cells");
                }
            }
        }
        return data;
    }

    public String[][] getCustomSheetForTDD() throws IOException {
        File file = new File(excelPath);
        FileInputStream fileInputStream = new FileInputStream(file);
        book = new XSSFWorkbook(fileInputStream);
        sheet = book.getSheet(sheetName);
        int rowNumber = xlsxRowCount();
        int columnNumber = xlsxColumnCount();
        String[][] data = new String[rowNumber - 1][columnNumber];
        for (int i = 1; i < rowNumber; i++) {
            for (int j = 0; j < columnNumber; j++) {
                XSSFRow row = sheet.getRow(i);
                XSSFCell cell = row.getCell(j);
                String value = cellToString(cell);
                data[i - 1][j] = value;
                if (value == null) {
                    System.out.println("Empty cells");
                }
            }
        }
        return data;
    }
}