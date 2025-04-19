import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.sql.Types;
import java.util.Date;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

@Getter
@Setter
public class ExcelService {
    /**
     * ApplyReadMethod
     *
     * @description Định nghĩa lambda function theo function interface
     */
    public interface ApplyReadMethod {
        void apply(int startRow, int endRow, Sheet sheet);

    }

    /**
     * ApplyWriteMethod
     *
     * @description Định nghĩa lambda function theo function interface
     */
    public interface ApplyWriteMethod {
        void apply(Sheet sheet);

    }

    @Setter
    @Getter
    @Builder
    public static class BuildStyle {
        private String fontName;
        private int fontSize;
        private int fontColor;
        private boolean fontBold;
        private boolean fontItalic;
        private boolean fontUnderline;
        private Workbook _workbook;


        public CellStyle getStyle() {
            CellStyle style = _workbook.createCellStyle();
            Font font = _workbook.createFont();

            // Định dạng Font
            font.setFontName(fontName);         // Đặt font chữ
            font.setFontHeightInPoints((short) fontSize); // Kích thước font
            font.setBold(fontBold);               // In đậm
            font.setItalic(fontItalic);             // In nghiêng
            if (fontUnderline) {
                font.setUnderline(Font.U_SINGLE); // Gạch dưới
            }
            font.setColor((short) fontColor); // Đặt màu đỏ

            // Gắn Font vào Style
            style.setFont(font);
            return style;
        }
    }

    private File file;
    private Workbook _workbook;
    private int _sheetCount;
    private boolean isFileValid;

    public static ExcelService createExcelFile(String fileName) {
        return new ExcelService("");
    }

    public File getFile() {
        return file;
    }

    public boolean deleteFile() {
        return file.delete();
    }

    /**
     * ExcelService
     *
     * @param filePath đường dẫn đến file
     * @description Khởi tạo dịch vụ cho file Excel
     */
    public ExcelService(String filePath) {
        try {
            file = new File(filePath);
            FileInputStream fis = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(fis);
            // Lấy số lượng sheet của file
            _sheetCount = workbook.getNumberOfSheets();
            _workbook = workbook;
            isFileValid = true;
        } catch (Exception e) {
            isFileValid = false;
            System.out.println("File not found");
        }
    }

    /**
     * readExcel
     *
     * @param sheetName   sheet muốn đọc
     * @param numOfThread số lượng thread muốn sử dụng
     * @param startRow    hàng bắt đâu đọc
     * @param endRow      hàng kết thúc
     * @param applyMethod lambda sẽ chạy
     * @description Hàm này đọc dữ liệu từ sheet theo thứ tự từ startRow đến endRow và chia vào thread để tối ưu hiệu năng đọc
     */
    public void readExcel(String sheetName, int numOfThread, int startRow, int endRow, ApplyReadMethod applyMethod) throws InterruptedException {
        long startTime = System.nanoTime();
        Sheet sheet = _workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new RuntimeException("Sheet not found: " + sheetName);
        }
        int sheetRowCount = sheet.getLastRowNum();

        if (endRow < startRow) {
            throw new IllegalArgumentException("Dong cuoi phai lon hon dong dau");
        } else {
            if (endRow > sheetRowCount) {
                endRow = sheetRowCount;
            }
            sheetRowCount = endRow;
        }

        startRow--;
        startRow = Math.max(startRow, 0);
        numOfThread = Math.max(numOfThread, 1);
        int actualRowCount = sheetRowCount - startRow;
        int rowPerThread = actualRowCount / numOfThread;
        rowPerThread = Math.max(rowPerThread, 1);
        int redundantRows = actualRowCount % numOfThread;
        redundantRows = rowPerThread == 1 ? 0 : redundantRows;
        int nextStartRow = 0;

        ExecutorService executorService = Executors.newFixedThreadPool(numOfThread);

        while (numOfThread > 0) {
            int start = startRow + nextStartRow;
            int end;
            if (redundantRows > 0) {
                end = start + rowPerThread + 1;
            } else {
                end = start + rowPerThread;
            }
            if (start > sheetRowCount) {
                break;
            }
            if (end > sheetRowCount) {
                end = sheetRowCount;
            }
            int finalEnd = end;
            executorService.submit(() -> {
                applyMethod.apply(start, finalEnd, sheet);
            });
            startRow = 0;
            nextStartRow = end + 1;
            numOfThread--;
        }


        executorService.shutdown();
        executorService.awaitTermination(1, TimeUnit.HOURS);
        long endTime = System.nanoTime();
        long duration = endTime - startTime;

        System.out.println("Thời gian chạy: " + (duration / 1_000_000) + " milli giây");
    }

    public void writeExcel(String sheetName, ApplyWriteMethod applyMethod) throws Exception {
        long startTime = System.nanoTime();
        Sheet sheet = _workbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = _workbook.createSheet(sheetName);
        }
        applyMethod.apply(sheet);

        FileOutputStream outputStream = new FileOutputStream(file.getAbsolutePath());
        _workbook.write(outputStream);

        long endTime = System.nanoTime();
        long duration = endTime - startTime;

        System.out.println("Thời gian chạy: " + (duration / 1_000_000) + " milli giây");
    }


    public static void setCellValue(Sheet sheet, int row, int cell, Object value, int type, BuildStyle buildStyle) {
        Row r = sheet.getRow(row);
        if (r == null) {
            r = sheet.createRow(row);
        }
        Cell c = r.getCell(cell);
        if (c == null) {
            c = r.createCell(cell);
        }
        switch (type) {
            case Types.INTEGER -> {
                c.setCellValue(Integer.parseInt(value.toString()));
            }
            case Types.DECIMAL -> {
                c.setCellValue(Double.parseDouble(value.toString()));
            }
            case Types.NVARCHAR -> {
                c.setCellValue((String) value);
            }
            default -> {
                c.setCellValue((String) value);
            }
        }
        if(buildStyle != null) {
            c.setCellStyle(buildStyle.getStyle());
        }else{
            c.setCellStyle(null);
        }
    }

    public static void setCellValue(Sheet sheet, int row, int cell, Date value, String format, BuildStyle buildStyle) {
        Row r = sheet.getRow(row);
        if (r == null) {
            r = sheet.createRow(row);
        }
        Cell c = r.getCell(cell);
        if (c == null) {
            c = r.createCell(cell);
        }
        c.setCellValue( value);
        if(buildStyle != null) {
            CellStyle cellStyle = buildStyle.getStyle();
            CreationHelper createHelper = sheet.getWorkbook().getCreationHelper();
            cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(format));
            c.setCellStyle(cellStyle);
        }else{
            CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
            CreationHelper createHelper = sheet.getWorkbook().getCreationHelper();
            cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(format));
            c.setCellStyle(buildStyle.getStyle());
        }
    }

}
