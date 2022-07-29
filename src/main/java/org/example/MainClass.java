package org.example;
import org.example.Excel_Action.Excel;
import org.example.Excel_Action.Excel_Action;
import org.example.action.read_properties_file;

import static org.example.Excel_Action.Excel.generate_excel_file;
import static org.example.setting.excel.filenameArrayList;
import static org.example.setting.filepath.default_work_book;

public class MainClass {
    public static void main(String[] args) {

        Excel.gen_excel_file();
        for (int counter = 0; counter < filenameArrayList.size(); counter++) {

            String file_name = read_properties_file.readfile(counter);
            Excel_Action.gen_sheet(file_name,counter);
        }
        generate_excel_file(default_work_book);


    }
}
