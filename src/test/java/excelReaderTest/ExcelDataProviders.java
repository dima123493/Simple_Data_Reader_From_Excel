package excelReaderTest;

import org.testng.annotations.DataProvider;

public class ExcelDataProviders {
    @DataProvider
    public Object[][] usersFromSheet1() throws Exception {
        String path = "src/main/resources/Data.xlsx";
        ExcelReader excelReader = new ExcelReader(path);
        return excelReader.getDataSheetForTDD();
    }
    @DataProvider
    public Object[][] usersFromSelectedSheet() throws Exception {
        String path = "src/main/resources/Data.xlsx";
        ExcelReader excelReader = new ExcelReader(path,"Лист2");
        return excelReader.getCustomSheetForTDD();
    }
}
