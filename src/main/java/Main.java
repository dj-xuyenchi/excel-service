import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.sql.Types;
import java.util.Date;

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
        String filePath2 = "output.xlsx";
        ExcelService e1 = new ExcelService(filePath2);
        boolean ss = e1.deleteFile();
        ExcelService e = new ExcelService(filePath);
        e.readExcel("Sheet1", 3, 1, 222, (x, s, sheet) -> {
            for (int i = x; i <= s; i++) {

                try {
                    Double d = sheet.getRow(i).getCell(0).getNumericCellValue();
                    System.out.println(d );
                    Thread.sleep(10);
                } catch (Exception es) {
                    System.out.println(es.getMessage());
                }
            }
        });

        e.writeExcel("gg3g", (d) -> {
            ExcelService.BuildStyle b = new ExcelService.BuildStyle.BuildStyleBuilder()._workbook(d.getWorkbook())
                    .fontSize(26)
                    .fontColor(IndexedColors.BLUE.getIndex())
                    .fontBold(true)
                    .fontItalic(true)
                    .fontUnderline(true)
                    .build()
                    ;
            for (int i = 1; i < 10; i++) {
                ExcelService.setCellValue(d, i, 1, "sdfsdfsd2f", Types.NVARCHAR,null);
            }
        });
        System.out.println("Ket thuc");


    }
}
