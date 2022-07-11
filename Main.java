import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class Main {

    public static void main(String[] args) throws IOException{
        createExcel();
    }


    // Пустой файл
    public static void createExcel() throws IOException{
        XSSFWorkbook book = new XSSFWorkbook();
        FileOutputStream file = new FileOutputStream("exp.xlsx");

        XSSFSheet sheet = book.createSheet("Sheet");

        book.write(file);
        file.close();
    }

    // Пустой файл с указанным именем
    public static void createExcel(String name) throws IOException{
        XSSFWorkbook book = new XSSFWorkbook();
        FileOutputStream file = new FileOutputStream(name + ".xlsx");

        XSSFSheet sheet = book.createSheet("Sheet");

        book.write(file);
        file.close();
    }

    // Файл с указанным именем и указанным значением
    public static void createExcel(String name, int cellNum, int rowNum, String value) throws IOException {
        XSSFWorkbook book = new XSSFWorkbook();
        FileOutputStream file = new FileOutputStream(name + ".xlsx");
        XSSFSheet sheet = book.createSheet("Sheet");

        Cell cell;
        Row row;

        row = sheet.createRow(rowNum);
        cell = row.createCell(cellNum, CellType.STRING);
        cell.setCellValue(value);

        book.write(file);
        file.close();
    }

    // Файл с указанным именем, указанным значением, шрифтом и стилем
    public static void createExcel(String name, int cellNum, int rowNum, String value, String fontName, boolean isItalic, boolean isBold) throws IOException {
        XSSFWorkbook book = new XSSFWorkbook();
        FileOutputStream file = new FileOutputStream(name + ".xlsx");
        XSSFSheet sheet = book.createSheet("Sheet");
        XSSFCellStyle style = book.createCellStyle();
        XSSFFont font = book.createFont();
        font.setFontName(fontName);
        font.setItalic(isItalic);
        font.setBold(isBold);
        style.setFont(font);

        Cell cell;
        Row row;

        row = sheet.createRow(rowNum);
        cell = row.createCell(cellNum, CellType.STRING);
        cell.setCellValue(value);
        cell.setCellStyle(style);

        book.write(file);
        file.close();
    }

    // Файл с указанным именем, указанным значением, шрифтом, и толщиной букв
    public static void createExcel(String name, int cellNum, int rowNum, String value, String fontName, boolean isItalic, boolean isBold, boolean isThin) throws IOException {
        XSSFWorkbook book = new XSSFWorkbook();
        FileOutputStream file = new FileOutputStream(name + ".xlsx");
        XSSFSheet sheet = book.createSheet("Sheet");
        XSSFCellStyle style = book.createCellStyle();
        XSSFFont font = book.createFont();
        font.setFontName(fontName);
        font.setItalic(isItalic);
        font.setBold(isBold);
        style.setFont(font);

        if(isThin) {
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
        }
        else{
            style.setBorderBottom(BorderStyle.THICK);
            style.setBorderTop(BorderStyle.THICK);
            style.setBorderLeft(BorderStyle.THICK);
            style.setBorderRight(BorderStyle.THICK);
        }

        Cell cell;
        Row row;

        row = sheet.createRow(rowNum);
        cell = row.createCell(cellNum, CellType.STRING);
        cell.setCellValue(value);
        cell.setCellStyle(style);

        book.write(file);
        file.close();
    }

    // Файл с указанным именем, указанным значением и расположением
    public static void createExcel(String name, int cellNum1, int rowNum1, int cellNum2, int rowNum2, String value, int allignmentH, int allignmentV) throws IOException {
        XSSFWorkbook book = new XSSFWorkbook();
        FileOutputStream file = new FileOutputStream(name + ".xlsx");
        XSSFSheet sheet = book.createSheet("Sheet");
        XSSFCellStyle style = book.createCellStyle();

        Cell cell;
        Row row;

        sheet.addMergedRegion(new CellRangeAddress(rowNum1,rowNum2,cellNum1,cellNum2));
        row = sheet.createRow(rowNum1);
        cell = row.createCell(cellNum1, CellType.STRING);
        cell.setCellValue(value);

        switch (allignmentH){
            case 1:
                style.setAlignment(HorizontalAlignment.LEFT);
                break;
            case 2:
                style.setAlignment(HorizontalAlignment.CENTER);
                break;
            case 3:
                style.setAlignment(HorizontalAlignment.RIGHT);
                break;
        }

        switch (allignmentV){
            case 1:
                style.setVerticalAlignment(VerticalAlignment.TOP);
                break;
            case 2:
                style.setVerticalAlignment(VerticalAlignment.CENTER);
                break;
            case 3:
                style.setVerticalAlignment(VerticalAlignment.BOTTOM);
                break;
        }

        cell.setCellStyle(style);

        book.write(file);
        file.close();
    }

    // Файл с указанным именем и заполнение значением
    public static void createExcel(String name, int cellNum1, int rowNum1, int cellNum2, int rowNum2, String value) throws IOException {
        XSSFWorkbook book = new XSSFWorkbook();
        FileOutputStream file = new FileOutputStream(name + ".xlsx");
        XSSFSheet sheet = book.createSheet("Page");

        Cell cell;
        Row row;

        for (; rowNum1 <= rowNum2; rowNum1++) {
            Row rt = sheet.createRow(rowNum1);
            for (int cellIndex = cellNum1; cellIndex <= cellNum2; cellIndex++) {
                rt.createCell(cellIndex).setCellValue(value);
            }
        }

        book.write(file);
        file.close();
    }
}
