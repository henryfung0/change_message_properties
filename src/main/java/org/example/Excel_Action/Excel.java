package org.example.Excel_Action;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.action.cmd_action;
import org.example.setting.filepath;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import static org.example.action.cmd_action.turtoise_check_for_modification;
import static org.example.setting.filepath.default_work_book;

public class Excel {
    public static void gen_excel_file() {
        String user_director = System.getProperty("user.dir") + filepath.excel_file;
        File f = new File(user_director);
        if (f.exists() && !f.isDirectory()) {
            //excel file exist
            System.out.println("excel file exist : true");
            try {
                FileInputStream file = new FileInputStream(user_director);
                //ZipSecureFile.setMinInflateRatio(-1.0d);
                XSSFWorkbook workbook = new XSSFWorkbook(file);
                Excel_Overwrite_Properties.check_modification(workbook);
                Excel_layout.preparelayout(workbook);
                int setting_position = -1;
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    if (workbook.getSheetName(i).equals("setting")) {
                        setting_position = i;
                        break;
                    }
                }
                remove_all_sheet_except_setting(workbook, setting_position);
                Excel_Setting.check_setting(workbook);
                default_work_book = workbook;
                turtoise_check_for_modification(filepath.SVN_path, filepath.project_path);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        } else {
            //excel file not exist
            System.out.println("excel file exist : false");
            XSSFWorkbook workbook = new XSSFWorkbook();
            Excel_layout.preparelayout(workbook);
            Excel_Setting.generate_setting_worksheet(workbook);
            generate_excel_file(workbook);
            System.exit(0);
        }
    }

    private static void remove_all_sheet_except_setting(XSSFWorkbook workbook, int setting_position) {
        int total_sheet_size = workbook.getNumberOfSheets();
        for (int sheet = total_sheet_size - 1; sheet >= 0; sheet--) {
            if (sheet == setting_position) {
                System.out.println("setting position : " + setting_position);
            } else {
                workbook.removeSheetAt(sheet);
            }
        }
    }

    public static String get_cell_value(XSSFSheet sheet, int col_num, int row_num) {
        Row row = sheet.getRow(row_num - 1);
        if (row != null) {
            Cell cell = row.getCell(col_num - 1);
            if (cell == null || cell.getCellType() == CellType.BLANK) {
                // This cell is empty
            } else {
                switch (cell.getCellType()) {
                    case NUMERIC -> {
                        return String.valueOf(cell.getNumericCellValue());
                    }
                    case STRING -> {
                        return cell.getStringCellValue();
                    }
                }
            }
        }
        return null;
    }

    public static void generate_excel_file(XSSFWorkbook workbook) {
        boolean normal = false;
        while (!normal) {
            try {
                normal=true;
                File file = new File(filepath.excel_sheet_name);
                FileOutputStream out = new FileOutputStream(file);
                workbook.write(out);
                out.close();
                System.out.println("message_properties.xlsx written successfully on disk.");
                cmd_action.open_excel(filepath.excel_sheet_name);
            } catch (Exception e) {
                normal=false;
                cmd_action.close_excel(filepath.excel_sheet_name);
            }
        }
    }




    public static String retrieve_cell_value(Cell cell) {
        return switch (cell.getCellType()) {
            case NUMERIC -> Integer.toString((int) cell.getNumericCellValue());
            case STRING -> cell.getStringCellValue();
            default -> null;
        };
    }




}
