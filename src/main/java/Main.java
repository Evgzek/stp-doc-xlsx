import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.io.FileOutputStream;

public class Main {
    public static void main(String[] args) {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("test");
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("Times New Roman");
        CellStyle cellStyle = wb.createCellStyle();
        CellStyle cellStyle1 = wb.createCellStyle();
        CellStyle cellStyle2 = wb.createCellStyle();
        CellStyle cellStyle3 = wb.createCellStyle();
        cellStyle3.setAlignment(HorizontalAlignment.CENTER);
        cellStyle3.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFont(font);
        cellStyle1.setFont(font);
        cellStyle.setWrapText(true);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle1.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle1.setAlignment(HorizontalAlignment.LEFT);
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(7);
        cell.setCellValue("УТВЕРЖАДЮ:");
        sheet.autoSizeColumn(7);
        row = sheet.createRow(1);
        cell = row.createCell(7);
        cell.setCellStyle(cellStyle2);
        cell.setCellValue("Директор");
        row = sheet.createRow(2);
        cell = row.createCell(8);
        cell.setCellStyle(cellStyle2);
        row.setHeightInPoints(35);
        cell.setCellValue("(сокращенное наименование\r\n образовательного учреждения)");
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 8, 9));
        sheet.setColumnWidth(8, 15*256);
        sheet.setColumnWidth(9, 15*256);
        cell.setCellStyle(cellStyle);
        row = sheet.createRow(3);
        cell = row.createCell(8);
        cell.setCellValue("_____________________________");
        row.setHeightInPoints(35);
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 8,9));
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 8,9));
        RegionUtil.setBorderTop(BorderStyle.DOTTED, new CellRangeAddress(2,2,8,9), sheet);
        cell = row.createCell(7);
        cell.setCellValue("______________");
        sheet.addMergedRegion(new CellRangeAddress(4,4,8,9));
        row = sheet.createRow(4);
        cell = row.createCell(7);
        cell.setCellValue("(подпись)");
        cell.setCellStyle(cellStyle);
        cell = row.createCell(8);
        cell.setCellValue("(расшифровка подписи)");
        cell.setCellStyle(cellStyle);
        row = sheet.createRow(6);
        cell = row.createCell(7);
        cell.setCellStyle(cellStyle1);
        cell.setCellValue("14.05.2022");
        row = sheet.createRow(7);
        cell = row.createCell(7);
        cell.setCellValue("М.П.");
        row = sheet.createRow(9);
        cell = row.createCell(0);
        sheet.addMergedRegion(new CellRangeAddress(9, 9, 0,9));
        cell.setCellValue("Отчет о фактическом предоставленном бесплатном питании");
        cell.setCellStyle(cellStyle3);
        row = sheet.createRow(10);
        cell = row.createCell(0);
        sheet.addMergedRegion(new CellRangeAddress(10, 10, 0,9));
        cell.setCellValue("за период с 01.05.2022 по 31.05.2022");
        cell.setCellStyle(cellStyle3);
        row = sheet.createRow(11);
        cell = row.createCell(0);
        RegionUtil.setBorderTop(BorderStyle.DASHED, new CellRangeAddress(11,11,0,6), sheet);
        cell.setCellStyle(cellStyle3);









        try {
            FileOutputStream out = new FileOutputStream("test.xls");
            wb.write(out);
            out.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
