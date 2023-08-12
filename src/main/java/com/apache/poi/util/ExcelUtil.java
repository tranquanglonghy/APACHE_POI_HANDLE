package com.apache.poi.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.tomcat.util.http.fileupload.ByteArrayOutputStream;
import org.springframework.stereotype.Component;

import java.io.ByteArrayInputStream;
import java.io.IOException;

@Component
public class ExcelUtil {
    private String SHEET_NAME = "Merge_cell_handle";

    private CellStyle getGlobalCellStyle(Workbook workbook) {
        CellStyle globalCellStyle = workbook.createCellStyle();
        globalCellStyle.setAlignment(HorizontalAlignment.CENTER);
        globalCellStyle.setBorderBottom(BorderStyle.THIN);
        globalCellStyle.setBorderTop(BorderStyle.THIN);
        globalCellStyle.setBorderRight(BorderStyle.THIN);
        globalCellStyle.setBorderLeft(BorderStyle.THIN);
        return globalCellStyle;
    }

    public ByteArrayInputStream tutorialsToExcel() {
        try (Workbook workbook = new SXSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet(SHEET_NAME);
            CellStyle cellStyle = getGlobalCellStyle(workbook);
            int rowindex1 = 1;
            int rowindex2 = 2;
            Row header1 = sheet.createRow(rowindex1);
            Row header2 = sheet.createRow(rowindex2);

            for (int i = 0; i < 5 * 3; i++) {
                Cell header1Cell = header1.createCell(i);
                Cell header2Cell = header2.createCell(i);
                header1Cell.setCellStyle(cellStyle);
                header2Cell.setCellStyle(cellStyle);
            }
            for (int i = 0; i < 3; i++) {
                Cell cell = header1.getCell(i * 5);
                cell.setCellValue("Week " + (i + 1) + " data");
                sheet.addMergedRegion(new CellRangeAddress(rowindex1, rowindex1, i * 5, i * 5 + 4));
            }

            for (int i = 0; i < 3; i++) {
                Cell cell1 = header2.getCell(i * 5);
                Cell cell2 = header2.getCell(i * 5 + 3);
                cell1.setCellValue("Week data 01");
                cell2.setCellValue("Week data 02");
                sheet.addMergedRegion(new CellRangeAddress(rowindex2, rowindex2, i * 5, i * 5 + 2));
                sheet.addMergedRegion(new CellRangeAddress(rowindex2, rowindex2, i * 5 + 3, i * 5 + 4));
            }

            for (int i = 0; i < 5; i++) {
                Row row = sheet.createRow(i + 3);
                for (int j = 0; j < 5 * 3; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue("data");
                }
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            throw new RuntimeException("fail to export data to Excel file: " + e.getMessage());
        }
    }
}
