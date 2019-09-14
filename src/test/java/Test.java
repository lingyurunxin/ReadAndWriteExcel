import com.liangpi.ReadExcel;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressBase;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Set;

/**
 * @author : 李远锋
 * @time : 2019-09-13 15:45
 * 文件说明：
 */

public class Test {


    public static void main(String[] args) throws IOException {

        ReadExcel readExcel = new ReadExcel();
        XSSFWorkbook workbook = new XSSFWorkbook("D:\\Java\\ReadAndExcel\\src\\main\\resources\\8.16.xlsx");
        XSSFSheet sheetAt = workbook.getSheetAt(0);
        readExcel.readOneSheet(sheetAt);
//        HashMap<String, HashMap<String, List<Double>>> allData = readExcel.getAllData();
//        System.out.println(allData);


    }

}
