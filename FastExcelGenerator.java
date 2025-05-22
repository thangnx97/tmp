package com.example.demo.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.util.Base64;

public class FastExcelGenerator {
    public static void main(String[] args) throws Exception {
        long startTime = System.currentTimeMillis();

        SXSSFWorkbook workbook = new SXSSFWorkbook(10000); // keep 5000 rows in memory, others are flushed to disk
        SXSSFSheet sheet = (SXSSFSheet) workbook.createSheet("Report");

        // Set fixed column widths
        for (int i = 0; i < 13; i++) {
            sheet.setColumnWidth(i, 20 * 256); // 20 characters wide
        }

        // Create styles
        CellStyle headerStyle = createCellStyle(workbook, HorizontalAlignment.CENTER, true);
        CellStyle centerStyle = createCellStyle(workbook, HorizontalAlignment.CENTER, false);
        CellStyle rightStyle = createCellStyle(workbook, HorizontalAlignment.RIGHT, false);
        CellStyle numberStyle = createNumberStyle(workbook);

        // Create header row
        Row header = sheet.createRow(0);
        header.setHeightInPoints(25); // Increase row height
        for (int i = 0; i < 13; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue("Header " + (i + 1));
            cell.setCellStyle(headerStyle);
        }

        // Generate 1 million rows
        for (int rowNum = 1; rowNum <= 1_000_000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            row.setHeightInPoints(20); // Set row height
            for (int col = 0; col < 13; col++) {
                Cell cell = row.createCell(col);
                if (col >= 9) {
                    cell.setCellValue(rowNum + col); // Example numeric data
                    cell.setCellStyle(numberStyle);
                } else {
                    cell.setCellValue("Dữ liệu dòng " + rowNum + "," + (col + 1));
                    cell.setCellStyle(col < 8 ? centerStyle : rightStyle);
                }
            }
            // Optionally flush rows to manage memory
            if (rowNum % 10000 == 0) {
                sheet.flushRows(5000); // retain 5000 rows in memory
            }
        }

//        // Write to file
//        try (FileOutputStream out = new FileOutputStream("report.xlsx")) {
//            workbook.write(out);
//        }

        // Write workbook to ByteArrayOutputStream in RAM
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        workbook.write(baos);
        workbook.dispose(); // important to clean temp files

        byte[] excelBytes = baos.toByteArray();

        long endTime = System.currentTimeMillis();
        System.out.println("Excel generated in " + (endTime - startTime) + " ms");

        // Optionally encode as Base64
        String base64Excel = Base64.getEncoder().encodeToString(excelBytes);

        System.out.println("Excel Base64 length: " + base64Excel.length());
        startTime = System.currentTimeMillis();
        // If you want, save it to file from byte array
        try (FileOutputStream fos = new FileOutputStream("report_sxssf_ram.xlsx")) {
            fos.write(excelBytes);
        }

        endTime = System.currentTimeMillis();
        System.out.println("B64 generation completed in " + (endTime - startTime) + " ms");
    }

    private static CellStyle createCellStyle(Workbook wb, HorizontalAlignment align, boolean bold) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(align);
        style.setVerticalAlignment(VerticalAlignment.CENTER); // Align text in middle of row

        // Thin borders
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        Font font = wb.createFont();
        font.setFontName("Arial");
        font.setBold(bold);
        style.setFont(font);

        return style;
    }

    private static CellStyle createNumberStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        // Thin borders
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        DataFormat format = wb.createDataFormat();
        style.setDataFormat(format.getFormat("#,##0")); // Number format with comma

        Font font = wb.createFont();
        font.setFontName("Arial");
        style.setFont(font);

        return style;
    }
}
