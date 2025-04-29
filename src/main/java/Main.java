import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.math.BigDecimal;
import java.sql.Types;
import java.util.*;

public class Main {
    private static void method(int start, int end, Sheet sheet, Sheet sheet2) {
        for (int i = start; i <= end; i++) {
            Row row = sheet.getRow(i);
            if (row != null && row.getCell(0) != null) {
                double d = row.getCell(0).getNumericCellValue();

                Row row2 = sheet2.createRow(i);
                Cell cell = row.createCell(0);
                cell.setCellValue(d);

            }
        }
    }

    public static void main(String[] args) throws Exception {
        String filePath = "data.xlsx"; // Đường dẫn đến file Excel
        String filePath2 = "Book1.xlsx";
        ExcelService e1 = new ExcelService(filePath2);
//        boolean ss = e1.deleteFile();
        ExcelService e = new ExcelService(filePath);
        //  List<String> res = Collections.synchronizedList(new ArrayList<>());
//        e.read(1, 3, 1, -1, (x, s, sheet) -> {
//            for (int i = x; i <= s; i++) {
//                try {
//                    res.add(ExcelService.getCellValue(sheet,i,0));
//                    //   System.out.println(ExcelService.getCellValue(sheet,i,0));
//                    // Thread.sleep(10);
//                } catch (Exception es) {
//                    System.out.println(es.getMessage());
//                }
//            }
//        });
        System.out.println();
        List<TestClass> data = new ArrayList<>();

        for (int i = 0; i < 10; i++) {
            TestClass a = new TestClass();
            a.setName("hihi" + i);
            a.setPrice(new BigDecimal(3433));
            a.setCreatedAt(new Date());
            data.add(a);
        }
        String[] col = new String[3];
        col[0] = "name";
        col[1] = "price";
        col[2] = "createdAt";
        File r = e1.writeExcel("data2", 2, 2, data, col, UUID.randomUUID().toString());
        System.out.println("Ket thuc");
    }
}
