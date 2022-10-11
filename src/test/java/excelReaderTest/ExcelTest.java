package excelReaderTest;

import org.testng.annotations.Test;

public class ExcelTest {

    @Test(dataProvider = "usersFromSheet1", dataProviderClass = ExcelDataProviders.class)//"data"
    void test01(String... data) {
        System.out.println("User " + data[1] + " has id " + data[0]);
    }

    @Test(dataProvider = "usersFromSelectedSheet", dataProviderClass = ExcelDataProviders.class)
    void test02(String... data) {
        System.out.println("Login " + data[0] + " has password " + data[1]);
    }
}

/*    @DataProvider
    Object[][] data() {
        return new Object[][]{
                {"Dmytro"},
                {"Vlad"},
                {"Oleh"}
        };
    }*/