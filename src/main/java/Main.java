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
        CellStyle cellStyle4 = wb.createCellStyle();
        CellStyle cellStyle5 = wb.createCellStyle();
        cellStyle5.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle5.setAlignment(HorizontalAlignment.CENTER);
        cellStyle5.setFont(font);
        cellStyle4.setVerticalAlignment(VerticalAlignment.TOP);
        cellStyle4.setAlignment(HorizontalAlignment.CENTER);
        cellStyle4.setFont(font);
        cellStyle3.setAlignment(HorizontalAlignment.CENTER);
        cellStyle3.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle3.setWrapText(true);
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
        sheet.setColumnWidth(0, 10*256);
        sheet.setColumnWidth(1, 15*256);
        sheet.setColumnWidth(2, 10*256);
        sheet.setColumnWidth(3, 30*256);
        sheet.setColumnWidth(4, 12*256);
        sheet.setColumnWidth(5, 12*256);
        sheet.setColumnWidth(6, 15*256);
        sheet.setColumnWidth(7, 15*256);
        cell.setCellStyle(cellStyle);
        row = sheet.createRow(15);
        row.setHeightInPoints(25);
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
        RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(11,11,0,6), sheet);
        sheet.addMergedRegion(new CellRangeAddress(11, 11, 0,9));
        row = sheet.createRow(12);
        cell = row.createCell(0);
        sheet.addMergedRegion(new CellRangeAddress(12, 12, 0,9));
        cell.setCellValue("(сокращенное наименование образовательного учреждения)");
        cell.setCellStyle(cellStyle4);
        row = sheet.createRow(14);
        cell = row.createCell(0);
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 0,0));
        cell.setCellValue("№ п/п");
        cell.setCellStyle(cellStyle3);
        row.setHeightInPoints(45);
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 1,1));
        cell = row.createCell(1);
        cell.setCellValue("№ счета");
        cell.setCellStyle(cellStyle3);
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 2,2));
        cell = row.createCell(2);
        cell.setCellStyle(cellStyle3);
        cell.setCellValue("Класс");
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 3,3));
        cell = row.createCell(3);
        cell.setCellStyle(cellStyle3);
        cell.setCellValue("Ф.И. ребенка");
        sheet.addMergedRegion(new CellRangeAddress(14, 14, 4,5));
        cell = row.createCell(4);
        cell.setCellStyle(cellStyle3);
        cell.setCellValue("Дни посещения");
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 6,6));
        cell = row.createCell(6);
        cell.setCellStyle(cellStyle3);
        cell.setCellValue("Остаток на \r\nначало месяца, \r\nруб.");
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 7,7));
        cell = row.createCell(7);
        cell.setCellStyle(cellStyle3);
        cell.setCellValue("Поступило в \r\nтекущем \r\nмесяце на\r\n питание, руб.");
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 8,8));
        cell = row.createCell(8);
        cell.setCellStyle(cellStyle3);
        cell.setCellValue("Израсходовано в \r\nтекущем \r\nмесяце на\r\n питание, руб.");
        sheet.addMergedRegion(new CellRangeAddress(14, 15, 9,9));
        cell = row.createCell(9);
        cell.setCellStyle(cellStyle3);
        cell.setCellValue("Остаток на \r\nконец месяца, \r\nруб.");
        row = sheet.createRow(15);
        row.setHeightInPoints(35);
        cell = row.createCell(4);
        cell.setCellStyle(cellStyle3);
        cell.setCellValue("плановые");
        cell = row.createCell(5);
        cell.setCellStyle(cellStyle3);
        cell.setCellValue("фактические");
        RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(14,15,0,9), sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(14,15,0,9), sheet);
        for (int i = 0; i < 10; i++) {
            RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(14,15,i,i), sheet);
        }
        RegionUtil.setBorderTop(BorderStyle.THIN, new CellRangeAddress(15,15,4,5), sheet);
        for (int i = 0; i < 50; i++) {
            row = sheet.createRow(16 + i);
            cell = row.createCell(0);
            cell.setCellStyle(cellStyle5);
            cell.setCellValue(i+1);
            RegionUtil.setBorderBottom(BorderStyle.THIN, new CellRangeAddress(16+i,16+i,0,0), sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(16+i,16+i,0,0), sheet);
        }













        try {
            FileOutputStream out = new FileOutputStream("test.xls");
            wb.write(out);
            out.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
