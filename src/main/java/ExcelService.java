import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

public class ExcelService {
    /**
     * ApplyMethod
     *
     * @description Định nghĩa lambda function theo function interface
     */
    public interface ApplyMethod {
        void apply(int startRow, int endRow, Sheet sheet);
    }

    private Workbook _workbook;
    private int _sheetCount;
    private int[] _sheetRowCount;
    private boolean isFileValid;

    public static ExcelService createExcelFile(String fileName) {
        return new ExcelService("");
    }

    /**
     * ExcelService
     *
     * @param filePath đường dẫn đến file
     * @description Khởi tạo dịch vụ cho file Excel
     */
    public ExcelService(String filePath) {
        try {
            FileInputStream fis = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(fis);
            // Lấy số lượng sheet của file
            _sheetCount = workbook.getNumberOfSheets();
            // Lấy dòng cuối cùng có dữ liệu của từng sheet
            _sheetRowCount = new int[_sheetCount];
            for (int i = 0; i < _sheetCount; i++) {
                _sheetRowCount[i] = workbook.getSheetAt(i).getLastRowNum();
            }
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
     * @param sheetTab    sheet muốn đọc
     * @param numOfThread số lượng thread muốn sử dụng
     * @param startRow    hàng bắt đâu đọc
     * @param endRow      hàng kết thúc
     * @param applyMethod lambda sẽ chạy
     * @description Hàm này đọc dữ liệu từ sheet theo thứ tự từ startRow đến endRow và chia vào thread để tối ưu hiệu năng đọc
     */
    public void readExcel(int sheetTab, int numOfThread, int startRow, int endRow, ApplyMethod applyMethod) throws InterruptedException {
        long startTime = System.nanoTime();
        if (sheetTab > _sheetCount || sheetTab < 0) {
            throw new IllegalArgumentException("File excel khong co sheet " + sheetTab);
        }
        Sheet sheet = _workbook.getSheetAt(sheetTab - 1);
        int sheetRowCount = _sheetRowCount[sheetTab - 1];

        if (endRow < startRow) {
            throw new IllegalArgumentException("Dong cuoi phai lon hon dong dau");
        } else {
            if (endRow > sheetRowCount) {
                endRow = sheetRowCount;
            }
            sheetRowCount = endRow;
        }


        ExecutorService executorService = Executors.newFixedThreadPool(numOfThread);
        startRow--;
        startRow = Math.max(startRow, 0);
        numOfThread = Math.max(numOfThread, 1);
        int actualRowCount = sheetRowCount - startRow;
        int rowPerThread = actualRowCount / numOfThread;
        rowPerThread = Math.max(rowPerThread, 1);
        int redundantRows = actualRowCount % numOfThread;
        redundantRows = rowPerThread == 1 ? 0 : redundantRows;
        int nextStartRow = 0;

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

    public boolean isFileValid() {
        return isFileValid;
    }

    public void setFileValid(boolean fileValid) {
        isFileValid = fileValid;
    }

    public Workbook get_workbook() {
        return _workbook;
    }

    public void set_workbook(Workbook _workbook) {
        this._workbook = _workbook;
    }

    public int get_sheetCount() {
        return _sheetCount;
    }

    public void set_sheetCount(int _sheetCount) {
        this._sheetCount = _sheetCount;
    }
}
