package task13;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;

public class ReadExcel {
    private static Map<String, Map<String, String>> data;


    public static void init() {
        if (data == null) {
            data = new HashMap<>();
            loadData();
        }
    }


    private static void loadData() {
        //Read from excel -> apache POI
        // map<key,value> ---> map<Tc_name,Map<Col_name,Col_value>>
        try {

            FileInputStream fileInputStream = new FileInputStream("src/test/resources/Task13data.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = workbook.getSheet("Sheet1");
            Row header = sheet.getRow(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                Map<String, String> colValues = new HashMap<>();
                String name = row.getCell(0).getStringCellValue();
                for (int j = 1; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    Cell colHeader = header.getCell(j);
                    colValues.put(colHeader.getStringCellValue(), cell.getStringCellValue());
                }
                data.put(name, colValues);
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static Map<String, String> getData(String Name) {
        return data.get(Name);
    }

}
