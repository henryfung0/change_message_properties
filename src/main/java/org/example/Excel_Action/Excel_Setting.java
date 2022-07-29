package org.example.Excel_Action;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.constructor.Setting_filename;
import org.example.setting.excel;
import org.example.setting.excel_setting;
import org.example.setting.filepath;

import static org.example.Excel_Action.Excel.get_cell_value;
import static org.example.Excel_Action.Excel_layout.hide_setting_column;
import static org.example.setting.excel.filenameArrayList;
import static org.example.setting.excel_setting.excel_lock_password;

public class Excel_Setting {
    public static void generate_setting_worksheet(XSSFWorkbook workbook) {

        XSSFSheet sheet = workbook.createSheet(excel_setting.excel_setting);
        sheet.protectSheet(excel_lock_password);
        hide_setting_column(sheet);

        sheet.setColumnWidth(0, 7000);
        sheet.setColumnWidth(1, 12000);
        sheet.setColumnWidth(2, 2000);
        sheet.setColumnWidth(3, 7000);
        sheet.setColumnWidth(4, 22000);
        Row row = sheet.createRow(0);


        cell_description(workbook, sheet, 1, row, "project path", excel_setting.color1);
        cell_allow_edit( 2, row);
        cell_description(workbook, sheet, 4, row, "File", excel_setting.color2);
        cell_description(workbook, sheet, 5, row, "File Path", excel_setting.color2);

        row = sheet.createRow(1);
        cell_description(workbook, sheet, 4, row, "create sheet name", excel_setting.color1);
        cell_allow_edit( 5, row);
        row = sheet.createRow(2);
        cell_description(workbook, sheet, 1, row, "function", excel_setting.color2);
        cell_description(workbook, sheet, 2, row, "(T/F)", excel_setting.color2);
        cell_description(workbook, sheet, 4, row, "eng", excel_setting.color1);
        cell_allow_edit( 5, row);

        row = sheet.createRow(3);
        cell_description(workbook, sheet, 1, row, "highlight_function", excel_setting.color1);
        cell_allow_edit( 2, row);
        cell_description(workbook, sheet, 4, row, "eng2", excel_setting.color1);
        cell_allow_edit( 5, row);
        row = sheet.createRow(4);
        cell_description(workbook, sheet, 1, row, "hide_column_function", excel_setting.color1);
        cell_allow_edit( 2, row);
        cell_description(workbook, sheet, 4, row, "zh_CN", excel_setting.color1);
        cell_allow_edit( 5, row);
        row = sheet.createRow(5);
        cell_description(workbook, sheet, 4, row, "zh_HK", excel_setting.color1);
        cell_allow_edit( 5, row);


        for (int count = 0; count <= excel.max_sheet; count++) {
            int skip_row = 6;
            row = sheet.createRow(7 + skip_row * count);
            cell_description(workbook, sheet, 4, row, "create sheet name", excel_setting.color1);
            cell_allow_edit( 5, row);
            row = sheet.createRow(8 + skip_row * count);
            cell_description(workbook, sheet, 4, row, "eng", excel_setting.color1);
            cell_allow_edit( 5, row);
            row = sheet.createRow(9 + skip_row * count);
            cell_description(workbook, sheet, 4, row, "eng2", excel_setting.color1);
            cell_allow_edit( 5, row);
            row = sheet.createRow(10 + skip_row * count);
            cell_description(workbook, sheet, 4, row, "zh_CN", excel_setting.color1);
            cell_allow_edit( 5, row);
            row = sheet.createRow(11 + skip_row * count);
            cell_description(workbook, sheet, 4, row, "zh_HK", excel_setting.color1);
            cell_allow_edit( 5, row);
        }
    }

    private static void cell_allow_edit(int col_num, Row row) {

        Cell cell = row.createCell(col_num - 1);
        cell.setCellStyle(Excel_layout.unlockedCellStyle);
    }

    private static void cell_description(XSSFWorkbook workbook, XSSFSheet sheet, int col_num, Row row, String VariableName, short color) {

        Cell cell = row.createCell(col_num - 1);
        cell.setCellValue(VariableName);
        highlightcell(workbook, cell, color);
    }

    private static void highlightcell(XSSFWorkbook workbook, Cell cell, short color) {

        CellStyle style = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        //headerFont.setColor(IndexedColors.WHITE.index);
        style.setFillForegroundColor(color);
        // and solid fill pattern produces solid grey cell fill
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(headerFont);
        cell.setCellStyle(style);
    }

    public static void check_setting(XSSFWorkbook workbook) {
        try {
            XSSFSheet sheet = workbook.getSheetAt(0);
            read_setting(sheet);
            generate_filelist(sheet);


        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private static void generate_filelist(XSSFSheet sheet) {

        for (int count = 0; count <= excel.max_sheet; count++) {
            Setting_filename setting_filename = new Setting_filename();
            int skip_row = 6;
            if (get_cell_value(sheet, excel.filelist_column, (skip_row * count + 2)) != null) {
                setting_filename.setSheet_name(get_cell_value(sheet, excel.filelist_column, (skip_row * count + 2)));
            }
            if (get_cell_value(sheet, excel.filelist_column, (skip_row * count + 3)) != null) {
                setting_filename.setFile_name(get_cell_value(sheet, excel.filelist_column, (skip_row * count + 3)));
            }
            if (get_cell_value(sheet, excel.filelist_column, (skip_row * count + 4)) != null) {
                setting_filename.setEng_File_name(get_cell_value(sheet, excel.filelist_column, (skip_row * count + 4)));
            }
            if (get_cell_value(sheet, excel.filelist_column, (skip_row * count + 5)) != null) {
                setting_filename.setZH_Ch_File_name(get_cell_value(sheet, excel.filelist_column, (skip_row * count + 5)));
            }
            if (get_cell_value(sheet, excel.filelist_column, (skip_row * count + 6)) != null) {
                setting_filename.setZH_Hk_File_name(get_cell_value(sheet, excel.filelist_column, (skip_row * count + 6)));
            }
            if (setting_filename.getSheet_name() != null && setting_filename.getFile_name() != null && setting_filename.getEng_File_name() != null && setting_filename.getZH_Ch_File_name() != null && setting_filename.getZH_Hk_File_name() != null) {
                filenameArrayList.add(setting_filename);
            }
        }
    }

    private static void read_setting(XSSFSheet sheet) {
        if (get_cell_value(sheet, excel.setting_column, 1) != null) {
            filepath.project_path = get_cell_value(sheet, excel.setting_column, 1);
        }
        if(get_cell_value(sheet, excel.setting_column, 4)!=null) {
            if (get_cell_value(sheet, excel.setting_column, 4).toLowerCase().contains("t")) {
                excel_setting.highlight_function = true;
            } else if (get_cell_value(sheet, excel.setting_column, 4).toLowerCase().contains("f")) {
                excel_setting.highlight_function = false;
            }
        }

        if(get_cell_value(sheet, excel.setting_column, 5)!=null) {
            if (get_cell_value(sheet, excel.setting_column, 5).toLowerCase().contains("t")) {
                excel_setting.hide_column_function = true;
            } else if (get_cell_value(sheet, excel.setting_column, 5).toLowerCase().contains("f")) {
                excel_setting.hide_column_function = false;
            }
        }


    }

}
