package org.example.action;
import org.example.setting.excel_overwrite_properties;
import org.example.setting.excel_setting;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

public class cmd_action {

    public static void close_excel(String filename) {
        try {
            String cmd = "taskkill /fi \"WindowTitle eq Microsoft Excel - " + filename + "\"";
            Runtime.getRuntime().exec(new String[]{"cmd", "/c", cmd});
            TimeUnit.SECONDS.sleep(5);
        } catch (IOException | InterruptedException e) {
            throw new RuntimeException(e);
        }
    }

    public static void open_excel(String filename) {
        try {
            Runtime.getRuntime().exec(new String[]{"cmd", "/c", filename});
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void turtoise_check_for_modification(String turtoise_filepath , String project_filepath){
        if(excel_setting.turtoise_checkformodification_function) {
            if (excel_overwrite_properties.modified) {
                try {
                    turtoise_filepath = filepath_add_quotation(turtoise_filepath);
                    project_filepath = filepath_add_quotation(project_filepath);
                    String cmd = "START " + turtoise_filepath + " TortoiseProc.exe /command:repostatus /path:" + project_filepath + " /closeonend:0";
                    Runtime.getRuntime().exec(new String[]{"cmd", "/c", cmd});
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }
        }
    }

    private static String filepath_add_quotation(String filepath){
        return "\""+filepath+"\"";
    }
}
