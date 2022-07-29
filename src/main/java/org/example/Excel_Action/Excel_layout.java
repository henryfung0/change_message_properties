package org.example.Excel_Action;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.setting.excel_setting;
import org.example.setting.filepath;

import static org.example.setting.excel.filenameArrayList;

public class Excel_layout {
    public static void edit_layout(XSSFWorkbook workbook, XSSFSheet sheet,int counter) {
        //edit column width
        for (int x = 0; x <= 9; x++) {
            sheet.setColumnWidth(x, 7750);
        }
        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 4000);
        //set first row title
        set_first_row(sheet,counter);
        sheet.createFreezePane(0, 1);
        //hide empty column
        hide_column(sheet);
    }
    public static XSSFSheet hide_setting_column(XSSFSheet sheet) {
        if(excel_setting.hide_column_function) {
            sheet.createFreezePane(5, 0);
        }
        return sheet;
    }
    public static XSSFSheet hide_column(XSSFSheet sheet) {
        if(excel_setting.hide_column_function) {
            sheet.createFreezePane(10, 1);
        }
        return sheet;
    }

    public static void set_first_row(XSSFSheet sheet,int counter) {


        String new_variable ="new ";

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("variable name");
        cell.setCellStyle(heading_style);

        cell = row.createCell(1);
        cell.setCellValue(new_variable+"variable name");
        cell.setCellStyle(heading_style);

        cell = row.createCell(2);
        cell.setCellValue("eng1");
        cell.setCellStyle(heading_style);

        cell = row.createCell(3);
        cell.setCellValue(filepath.project_path+filenameArrayList.get(counter).getFile_name());
        cell.setCellStyle(heading_style);

        cell = row.createCell(4);
        cell.setCellValue("eng2");
        cell.setCellStyle(heading_style);

        cell = row.createCell(5);
        cell.setCellValue(filepath.project_path+filenameArrayList.get(counter).getEng_File_name());
        cell.setCellStyle(heading_style);

        cell = row.createCell(6);
        cell.setCellValue("Zh_hk");
        cell.setCellStyle(heading_style);

        cell = row.createCell(7);
        cell.setCellValue(filepath.project_path+filenameArrayList.get(counter).getZH_Hk_File_name());
        cell.setCellStyle(heading_style);

        cell = row.createCell(8);
        cell.setCellValue("Zh_ch");
        cell.setCellStyle(heading_style);

        cell = row.createCell(9);
        cell.setCellValue(filepath.project_path+filenameArrayList.get(counter).getZH_Ch_File_name());
        cell.setCellStyle(heading_style);
    }
    public static CellStyle unlockedCellStyle;
    public static CellStyle heading_style;

    public static CellStyle emptyfield_style;

    public static CellStyle WrapText_style;
    public static void preparelayout(XSSFWorkbook workbook){
        unlockedCellStyle = workbook.createCellStyle();
        unlockedCellStyle.setLocked(false);
        
        heading_style = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setColor(IndexedColors.WHITE.index);
        heading_style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
        // and solid fill pattern produces solid grey cell fill
        heading_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        heading_style.setFont(headerFont);

        emptyfield_style = workbook.createCellStyle();
        emptyfield_style.setFillForegroundColor(IndexedColors.YELLOW.index);
        emptyfield_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        emptyfield_style.setFont(headerFont);

        WrapText_style = workbook.createCellStyle();
        Excel_layout.WrapText_style.setWrapText(true);


    }
    public static void cell_allow_edit(Cell cell) {
        cell.setCellStyle(unlockedCellStyle);
    }
}
