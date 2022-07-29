package org.example.Excel_Action;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.action.read_properties_file;
import org.example.constructor.Change_Message;
import org.example.setting.excel_overwrite_properties;

import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.example.action.convert_words.StringToUnicode;
import static org.example.action.convert_words.convert_regrex;
import static org.example.action.read_properties_file.readvariable;
import static org.example.setting.excel.filenameArrayList;
import static org.example.setting.excel_overwrite_properties.change_messageArrayList;
import static org.example.setting.filepath.*;

public class Excel_Overwrite_Properties {
    public static void check_modification(XSSFWorkbook workbook) {
        for (int i = 1; i < workbook.getNumberOfSheets(); i++) {
            get_modify_variable(workbook.getSheetAt(i));
            replace_properties(excel_overwrite_properties.eng_filepath, eng_index);
            replace_properties(excel_overwrite_properties.eng2_filepath, eng2_index);
            replace_properties(excel_overwrite_properties.Zh_hk_filepath, zh_HK_index);
            replace_properties(excel_overwrite_properties.Zh_ch_filepath, zh_CN_index);
        }
    }

    private static void replace_properties(String filepath, int lang) {
        try {
            Path path = Paths.get(filepath);
            Charset charset = StandardCharsets.UTF_8;
            String content = Files.readString(path, charset);
            for (Change_Message change_message : change_messageArrayList) {
                String oldmessage = "";
                String newmessage = "";
                if (change_message.getNew_variable_name() == null) {
                    change_message.setNew_variable_name(change_message.getVariable_name());
                    excel_overwrite_properties.modified = true;
                }
                if (lang == eng_index) {
                    if (change_message.getNew_eng() != null) {
                        oldmessage = change_message.getVariable_name() + "=" + change_message.getEng();
                        newmessage = change_message.getNew_variable_name() + "=" + change_message.getNew_eng();
                        content = content.replaceAll(oldmessage, newmessage);
                        excel_overwrite_properties.modified = true;
                    }
                } else if (lang == eng2_index) {
                    if (change_message.getNew_eng2() != null) {
                        oldmessage = change_message.getVariable_name() + "=" + change_message.getEng2();
                        newmessage = change_message.getNew_variable_name() + "=" + change_message.getNew_eng2();
                        content = content.replaceAll(oldmessage, newmessage);
                        excel_overwrite_properties.modified = true;
                    }
                } else if (lang == zh_CN_index) {
                    if (change_message.getNew_zh_ch() != null) {
                        oldmessage = "(?i)" + change_message.getVariable_name() + "=" + convert_regrex(StringToUnicode(change_message.getZh_ch()));
                        newmessage = change_message.getNew_variable_name() + "=" + convert_regrex(StringToUnicode(change_message.getNew_zh_ch()));
                        content = content.replaceAll(oldmessage, newmessage);
                        excel_overwrite_properties.modified = true;
                    }
                } else if (lang == zh_HK_index) {
                    if (change_message.getNew_zh_hk() != null) {
                        oldmessage = "(?i)" + change_message.getVariable_name() + "=" + convert_regrex(StringToUnicode(change_message.getZh_hk()));
                        newmessage = change_message.getNew_variable_name() + "=" + convert_regrex(StringToUnicode(change_message.getNew_zh_hk()));
                        content = content.replaceAll(oldmessage, newmessage);
                        excel_overwrite_properties.modified = true;
                    }
                }
                Files.writeString(path, content, charset);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static String replaceall(String input, String regex, String replacement) {
        Pattern p = Pattern.compile(regex, Pattern.CASE_INSENSITIVE);
        Matcher m = p.matcher(input);
        String result = m.replaceAll(replacement);
        return result;
    }

    private static void get_modify_variable(Sheet sheet) {
        getcolumnvalue(sheet, 1, 3, 5, 7, 9);
    }

    private static void getcolumnvalue(Sheet sheet, int variable_name, int eng, int eng2, int Zh_hk, int Zh_ch) {
        Iterator<Row> rowIterator = sheet.iterator();
        String variable_name_filepath = "";
        boolean heading = true;
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            String variable_name_cell = Excel.retrieve_cell_value(row.getCell(variable_name - 1));
            String eng_cell = Excel.retrieve_cell_value(row.getCell(eng - 1));
            String eng2_cell = Excel.retrieve_cell_value(row.getCell(eng2 - 1));
            String Zh_hk_cell = Excel.retrieve_cell_value(row.getCell(Zh_hk - 1));
            String Zh_ch_cell = Excel.retrieve_cell_value(row.getCell(Zh_ch - 1));
            String new_variable_name_cell = Excel.retrieve_cell_value(row.getCell(variable_name));
            String new_eng_cell = Excel.retrieve_cell_value(row.getCell(eng));
            String new_eng2_cell = Excel.retrieve_cell_value(row.getCell(eng2));
            String new_Zh_hk_cell = Excel.retrieve_cell_value(row.getCell(Zh_hk));
            String new_Zh_ch_cell = Excel.retrieve_cell_value(row.getCell(Zh_ch));
            if (heading) {
                heading = false;
                excel_overwrite_properties.eng_filepath = new_eng_cell;
                excel_overwrite_properties.eng2_filepath = new_eng2_cell;
                excel_overwrite_properties.Zh_hk_filepath = new_Zh_hk_cell;
                excel_overwrite_properties.Zh_ch_filepath = new_Zh_ch_cell;
            } else {
                Change_Message change_message = new Change_Message();
                assert new_variable_name_cell != null;
                assert new_eng_cell != null;
                assert new_eng2_cell != null;
                assert new_Zh_hk_cell != null;
                assert new_Zh_ch_cell != null;
                if (checkEmpty(new_variable_name_cell) || checkEmpty(new_eng_cell) || checkEmpty(new_eng2_cell) || checkEmpty(new_Zh_hk_cell) || checkEmpty(new_Zh_ch_cell)) {
                    if ((!variable_name_cell.equals(new_variable_name_cell)) || (!eng_cell.equals(new_eng_cell)) || (!eng2_cell.equals(new_eng2_cell)) || (!Zh_hk_cell.equals(new_Zh_hk_cell)) || (!Zh_ch_cell.equals(new_Zh_ch_cell))) {
                        change_message.setVariable_name(variable_name_cell);
                        change_message.setNew_variable_name(new_variable_name_cell);
                        change_message.setEng(eng_cell);
                        change_message.setNew_eng(new_eng_cell);
                        change_message.setEng2(eng2_cell);
                        change_message.setNew_eng2(new_eng2_cell);
                        change_message.setZh_hk(Zh_hk_cell);
                        change_message.setNew_zh_hk(new_Zh_hk_cell);
                        change_message.setZh_ch(Zh_ch_cell);
                        change_message.setNew_zh_ch(new_Zh_ch_cell);
                        change_messageArrayList.add(change_message);
                    }
                }
            }
        }
    }

    private static boolean checkEmpty(String string) {
        return !(string == null || string.isEmpty());
    }

    public static String readfile(int counter) {
        readvariable(read_properties_file.format_filepath(filenameArrayList.get(counter).getFile_name()), eng_index);
        readvariable(read_properties_file.format_filepath(filenameArrayList.get(counter).getEng_File_name()), eng2_index);
        readvariable(read_properties_file.format_filepath(filenameArrayList.get(counter).getZH_Ch_File_name()), zh_CN_index);
        readvariable(read_properties_file.format_filepath(filenameArrayList.get(counter).getZH_Hk_File_name()), zh_HK_index);
        return filenameArrayList.get(counter).getSheet_name();
    }
}
