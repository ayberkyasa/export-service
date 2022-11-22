package com.msu.tt.claim.util;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class Export {

    /**
     * This method is to export a form element to excel file.
     * @param map
     * @param httpServletResponse
     * @throws IOException
     */
    public static void exportToExcel(HashMap<String, Object> map, HttpServletResponse httpServletResponse) throws IOException {
        cleanNulls(map);

        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("sheet1");

        int rowNumber = 0;

        for(HashMap.Entry<String, Object> entry : map.entrySet()) {
            XSSFRow row = sheet.createRow(rowNumber++);
            row.createCell(0).setCellValue(entry.getKey());
            row.createCell(1).setCellValue(String.valueOf(entry.getValue()));
        }

        httpServletResponse.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        httpServletResponse.setHeader("Content-disposition", "attachment;filename=" + "\"a.xlsx\"");
        try (OutputStream outputStream = httpServletResponse.getOutputStream()) {
            workbook.write(outputStream);
        }
    }

    /**
     * This method is to export a datatable to excel file.
     * @param mapList
     * @param httpServletResponse
     * @throws IOException
     */
    public static void exportToExcel(List<HashMap<String, Object>> mapList, HttpServletResponse httpServletResponse) throws IOException {
        cleanNulls(mapList);

        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("sheet1");

        int rowNumber = 0;
        XSSFRow row = sheet.createRow(rowNumber++);
        List<String> keys = new ArrayList<>(mapList.get(0).keySet());
        for (int i = 0; i < keys.size(); i++) {
            row.createCell(i).setCellValue(keys.get(i)); // add header row
        }

        for (HashMap<String, Object> map : mapList) {
            row = sheet.createRow(rowNumber++);
            for (int i = 0; i < keys.size(); i++) {
                row.createCell(i).setCellValue(String.valueOf(map.get(keys.get(i)))); // add data row
            }
        }

        httpServletResponse.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        httpServletResponse.setHeader("Content-disposition", "attachment;filename=" + "\"a.xlsx\"");
        try (OutputStream outputStream = httpServletResponse.getOutputStream()) {
            workbook.write(outputStream);
        }
    }

    /**
     * This method is to export multiple datatables to excel file.
     * @param mapList
     * @param httpServletResponse
     * @throws IOException
     */
    public static void exportMultipleDatatableToExcel(List<List<HashMap<String, Object>>> mapList, HttpServletResponse httpServletResponse) throws IOException {
        cleanNullsMultiple(mapList);

        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("sheet1");

        int rowNumber = 0;
        for(List<HashMap<String, Object>> hashMapList: mapList) {
            XSSFRow row = sheet.createRow(rowNumber++);
            List<String> keys = new ArrayList<>(hashMapList.get(0).keySet());
            for (int i = 0; i < keys.size(); i++) {
                row.createCell(i).setCellValue(keys.get(i)); // add header row
            }

            for (HashMap<String, Object> map : hashMapList) {
                row = sheet.createRow(rowNumber++);
                for (int i = 0; i < keys.size(); i++) {
                    row.createCell(i).setCellValue(String.valueOf(map.get(keys.get(i)))); // add data row
                }
            }
            row = sheet.createRow(rowNumber++);
        }

        httpServletResponse.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        httpServletResponse.setHeader("Content-disposition", "attachment;filename=" + "\"a.xlsx\"");
        try (OutputStream outputStream = httpServletResponse.getOutputStream()) {
            workbook.write(outputStream);
        }
    }

    /**
     *
     * @param map
     * @param httpServletResponse
     * @throws IOException
     * @throws DocumentException
     */
    public static void exportToPdf(HashMap<String, Object> map, HttpServletResponse httpServletResponse) throws IOException, DocumentException {
        BaseFont STF_Helvetica_Turkish = BaseFont.createFont("Helvetica", "CP1254", BaseFont.NOT_EMBEDDED);
        Font fontNormal = new Font(STF_Helvetica_Turkish, 12, Font.NORMAL);

        cleanNulls(map);

        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("sheet1");

        int rowNumber = 0;

        for(HashMap.Entry<String, Object> entry : map.entrySet()) {
            XSSFRow row = sheet.createRow(rowNumber++);
            row.createCell(0).setCellValue(entry.getKey());
            row.createCell(1).setCellValue(String.valueOf(entry.getValue()));
        }

        Document pdf = new Document();
        PdfWriter.getInstance(pdf, httpServletResponse.getOutputStream());
        pdf.open();
        PdfPTable table = new PdfPTable(2);
        PdfPCell cell;

        for (int i = 0; i < rowNumber; i++) {
            XSSFRow currentRow = sheet.getRow(i);
            cell=new PdfPCell(new Phrase(String.valueOf(currentRow.getCell(0)), fontNormal));
            table.addCell(cell);
            cell=new PdfPCell(new Phrase(String.valueOf(currentRow.getCell(1)), fontNormal));
            table.addCell(cell);
        }
        pdf.add(table);
        pdf.close();

        httpServletResponse.setContentType("application/pdf;charset=UTF-8");
        httpServletResponse.setHeader("Content-disposition", "attachment;filename=" + "\"a.pdf\"");
    }

    public static void exportToPdf(List<HashMap<String, Object>> mapList, HttpServletResponse httpServletResponse) throws IOException, DocumentException {
        BaseFont STF_Helvetica_Turkish = BaseFont.createFont("Helvetica", "CP1254", BaseFont.NOT_EMBEDDED);
        Font fontNormal = new Font(STF_Helvetica_Turkish, 12, Font.NORMAL);

        cleanNulls(mapList);

        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("sheet1");

        int rowNumber = 0;
        XSSFRow row = sheet.createRow(rowNumber++);
        List<String> keys = new ArrayList<>(mapList.get(0).keySet());
        for (int i = 0; i < keys.size(); i++) {
            row.createCell(i).setCellValue(keys.get(i)); // add header row
        }

        for (HashMap<String, Object> map : mapList) {
            row = sheet.createRow(rowNumber++);
            for (int i = 0; i < keys.size(); i++) {
                row.createCell(i).setCellValue(String.valueOf(map.get(keys.get(i)))); // add data row
            }
        }

        Document pdf = new Document();
        PdfWriter.getInstance(pdf, httpServletResponse.getOutputStream());
        pdf.open();
        PdfPTable table = new PdfPTable(keys.size());
        PdfPCell cell;

        for (int i = 0; i < rowNumber; i++) {
            XSSFRow currentRow = sheet.getRow(i);
            int lastColumn = currentRow.getLastCellNum();
            for(int j = 0; j < lastColumn; j++) {
                cell=new PdfPCell(new Phrase(String.valueOf(currentRow.getCell(j)), fontNormal));
                table.addCell(cell);
            }
        }
        pdf.add(table);
        pdf.close();

        httpServletResponse.setContentType("application/pdf;charset=UTF-8");
        httpServletResponse.setHeader("Content-disposition", "attachment;filename=" + "\"a.pdf\"");
    }

    private static void cleanNulls(HashMap<String, Object> map) {
        for(HashMap.Entry<String, Object> entry : map.entrySet()) {
            if(entry.getValue() == null)
                entry.setValue("");
        }
    }
    private static void cleanNulls(List<HashMap<String, Object>> mapList) {
        for(HashMap<String, Object> map: mapList) {
            for(HashMap.Entry<String, Object> entry : map.entrySet()) {
                if(entry.getValue() == null)
                    entry.setValue("");
            }
        }
    }
    private static void cleanNullsMultiple(List<List<HashMap<String, Object>>> mapList) {
        for(List<HashMap<String, Object>> hashMapList: mapList) {
            for(HashMap<String, Object> map: hashMapList) {
                for(HashMap.Entry<String, Object> entry : map.entrySet()) {
                    if(entry.getValue() == null)
                        entry.setValue("");
                }
            }
        }

    }
}
