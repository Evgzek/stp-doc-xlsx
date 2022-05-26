import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.io.FileOutputStream;

public class Main {
    public static void main(String[] args) {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("test");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setWrapText(true);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(7);
        cell.setCellValue("УТВЕРЖАДЮ:");
        sheet.autoSizeColumn(7);
        row = sheet.createRow(1);
        cell = row.createCell(7);
        cell.setCellValue("Директор");
        row = sheet.createRow(2);
        cell = row.createCell(8);
        row.setHeightInPoints(35);
        cell.setCellValue("(сокращенное наименование\r\n образовательного учреждения)");
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 8, 9));
        sheet.setColumnWidth(8, 15*256);
        sheet.setColumnWidth(9, 15*256);
        cell.setCellStyle(cellStyle);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 8,9));
        RegionUtil.setBorderTop(BorderStyle.DOTTED, new CellRangeAddress(2,2,8,9), sheet);
        try {
            FileOutputStream out = new FileOutputStream("test.xls");
            wb.write(out);
            out.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
