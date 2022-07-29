package org.example.action;
import org.example.constructor.Message;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;

import static org.example.action.convert_words.unicodeToUtf8;
import static org.example.setting.excel.filenameArrayList;
import static org.example.setting.filepath.*;

public class read_properties_file {
    public static ArrayList<Message> messageArrayList = new ArrayList<Message>(); // Create an ArrayList object

    public static String readfile(int counter) {
        reset_message_arraylist();
        readvariable(format_filepath(filenameArrayList.get(counter).getFile_name()), eng_index);
        readvariable(format_filepath(filenameArrayList.get(counter).getEng_File_name()), eng2_index);
        readvariable(format_filepath(filenameArrayList.get(counter).getZH_Ch_File_name()), zh_CN_index);
        readvariable(format_filepath(filenameArrayList.get(counter).getZH_Hk_File_name()), zh_HK_index);
        System.out.println("file name is :" + filenameArrayList.get(counter).getSheet_name());
        return filenameArrayList.get(counter).getSheet_name();
    }

    private static void reset_message_arraylist() {
        messageArrayList.clear();
    }

    public static String format_filepath(String file) {
        if (!file.toLowerCase().startsWith(project_path.toLowerCase())) {
            file = project_path + file;
        }
        if (!file.toLowerCase().endsWith(file_format.toLowerCase())) {
            file = file + file_format;
        }
        System.out.println(file);
        return file;
    }

    public static void readvariable(String file_name, int language) {
        try {
            File file = new File(file_name);
            BufferedReader br = new BufferedReader(new FileReader(file));
            String st;
            while ((st = br.readLine()) != null) {
                //check not comment
                if (!st.startsWith("#")) {
                    String[] split = st.split("=");
                    if (split.length > 1) {
                        //check have variable name and variable message
                        String variable = split[0];
                        String words = split[1];
                        int position = getarraylistpoistion(variable);
                        if (position == -1) {
                            //no record
                            Message message = new Message();
                            message.setVariable_name(variable);
                            message.setEng(words);
                            messageArrayList.add(message);
                        } else {
                            Message message = messageArrayList.get(position);
                            if (language == eng_index) {
                                message.setEng(words);
                            } else if (language == eng2_index) {
                                message.setEng2(words);
                            } else if (language == zh_CN_index) {
                                message.setZh_ch(unicodeToUtf8(words));
                            } else if (language == zh_HK_index) {
                                message.setZh_hk(unicodeToUtf8(words));
                            }
                            messageArrayList.set(position, message);
                        }
                    }
                }
            }
        } catch (IOException e) {
            System.out.println(e);
        }
    }

    public static int getarraylistpoistion(String variablename) {
        int position = -1;
        for (int i = 0; i < messageArrayList.size(); i++) {
            if (messageArrayList.get(i).getVariable_name().equals(variablename)) {
                position = i;
                break;
            }
        }
        return position;
    }

    private static void readrecord() {
        for (Message messageArrayList : messageArrayList) {
            System.out.print(messageArrayList.getVariable_name());
            System.out.print(":");
            System.out.println(messageArrayList.getEng());
        }
    }
}
