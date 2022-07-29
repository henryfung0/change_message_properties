package org.example.Excel_Action;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.constructor.Message;
import org.example.setting.excel_setting;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetProtection;

import static org.example.Excel_Action.Excel_layout.edit_layout;
import static org.example.action.read_properties_file.messageArrayList;
import static org.example.setting.excel_setting.excel_lock_password;
import static org.example.setting.filepath.default_work_book;

//import statements
public class Excel_Action {
    public static void gen_sheet(String file_name,int counter) {
        XSSFWorkbook workbook = default_work_book;
        XSSFSheet sheet = workbook.createSheet(file_name);

        int rownum = 1;
        sheet.protectSheet(excel_lock_password);
        for (Message messageArrayList : messageArrayList) {
            Row row = sheet.createRow(rownum++);
            Cell cell = row.createCell(0);
            cell.setCellValue(messageArrayList.getVariable_name());
            cell_style(cell, messageArrayList.getVariable_name());
            cell = row.createCell(1);
            org.example.Excel_Action.Excel_layout.cell_allow_edit(cell);
            cell = row.createCell(2);
            cell.setCellValue(messageArrayList.getEng());
            cell_style(cell, messageArrayList.getEng());
            cell = row.createCell(3);
            org.example.Excel_Action.Excel_layout.cell_allow_edit(cell);
            cell = row.createCell(4);
            cell.setCellValue(messageArrayList.getEng2());
            cell_style(cell, messageArrayList.getEng2());
            cell = row.createCell(5);
            org.example.Excel_Action.Excel_layout.cell_allow_edit(cell);
            cell = row.createCell(6);
            cell.setCellValue(messageArrayList.getZh_hk());
            cell_style(cell, messageArrayList.getZh_hk());
            cell = row.createCell(7);
            org.example.Excel_Action.Excel_layout.cell_allow_edit(cell);
            cell = row.createCell(8);
            cell.setCellValue(messageArrayList.getZh_ch());
            cell_style(cell, messageArrayList.getZh_ch());
            cell = row.createCell(9);
            org.example.Excel_Action.Excel_layout.cell_allow_edit(cell);
        }
        edit_layout(workbook, sheet,counter);
        //lockAll(sheet);
    }

    private static void lockAll(XSSFSheet xssfSheet) {
        xssfSheet.enableLocking();
        CTSheetProtection sheetProtection = xssfSheet.getCTWorksheet().getSheetProtection();
        sheetProtection.setSelectLockedCells(true);
        sheetProtection.setSelectUnlockedCells(false);
        sheetProtection.setFormatCells(true);
        sheetProtection.setFormatColumns(true);
        sheetProtection.setFormatRows(true);
        sheetProtection.setInsertColumns(true);
        sheetProtection.setInsertRows(true);
        sheetProtection.setInsertHyperlinks(true);
        sheetProtection.setDeleteColumns(true);
        sheetProtection.setDeleteRows(true);
        sheetProtection.setSort(false);
        sheetProtection.setAutoFilter(false);
        sheetProtection.setPivotTables(true);
        sheetProtection.setObjects(true);
        sheetProtection.setScenarios(true);
        ;
    }

    private static void cell_style(Cell cell, String getstring) {
        if (excel_setting.allWrapText_function) {
            cell.setCellStyle(Excel_layout.WrapText_style);
        }
        if (excel_setting.highlight_function) {

            if (getstring == null) {
                cell.setCellStyle(Excel_layout.emptyfield_style);
            }
        }
    }
}