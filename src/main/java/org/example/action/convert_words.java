package org.example.action;
import com.github.houbb.opencc4j.util.ZhConverterUtil;

import java.io.UnsupportedEncodingException;

public class convert_words {
    /*------------------------convert chinese chararcter to traditional or simplified ------------------------*/
    //https://github.com/houbb/opencc4j
    public static String traditional_chinese_to_simplified_chinese(String traditional_chinese) {
        String simplified_chinese = ZhConverterUtil.toSimple(traditional_chinese);
        return simplified_chinese;
    }

    public static String simplified_chinese_to_traditional_chinese(String simplified_chinese) {
        String traditional_chinese = ZhConverterUtil.toTraditional(simplified_chinese);
        return traditional_chinese;
    }

    /*------------------------convert chararcter to UTF-8 or UNICODE ------------------------*/
    public static String testing = "电子成立公司流程";


    public static String StringToUnicode(String s) {
        try {
            StringBuffer out = new StringBuffer("");
            byte[] bytes = s.getBytes("unicode");
            for (int i = 0; i < bytes.length - 1; i += 2) {
                out.append("\\u");
                String str = Integer.toHexString(bytes[i + 1] & 0xff);
                String str1 = Integer.toHexString(bytes[i] & 0xff);
                out.append(str1);
                for (int j = str.length(); j < 2; j++) {
                    out.append("0");
                }
                out.append(str);
            }
            return out.toString().substring(6);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
            return null;
        }
    }
    public static String convert_regrex(String regrex){
        return regrex.replaceAll("\\\\u","\\\\\\\\u");
    }
    public static String unicodeToUtf8(String theString) {
        char aChar;
        int len = theString.length();
        StringBuffer outBuffer = new StringBuffer(len);
        for (int x = 0; x < len; ) {
            aChar = theString.charAt(x++);
            if (aChar == '\\') {
                aChar = theString.charAt(x++);
                if (aChar == 'u') {
                    // Read the xxxx
                    int value = 0;
                    for (int i = 0; i < 4; i++) {
                        aChar = theString.charAt(x++);
                        switch (aChar) {
                            case '0':
                            case '1':
                            case '2':
                            case '3':
                            case '4':
                            case '5':
                            case '6':
                            case '7':
                            case '8':
                            case '9':
                                value = (value << 4) + aChar - '0';
                                break;
                            case 'a':
                            case 'b':
                            case 'c':
                            case 'd':
                            case 'e':
                            case 'f':
                                value = (value << 4) + 10 + aChar - 'a';
                                break;
                            case 'A':
                            case 'B':
                            case 'C':
                            case 'D':
                            case 'E':
                            case 'F':
                                value = (value << 4) + 10 + aChar - 'A';
                                break;
                            default:
                                throw new IllegalArgumentException("Malformed    \\uxxxx   encoding.");
                        }
                    }
                    outBuffer.append((char) value);
                } else {
                    if (aChar == 't') aChar = '\t';
                    else if (aChar == 'r') aChar = '\r';
                    else if (aChar == 'n') aChar = '\n';
                    else if (aChar == 'f') aChar = '\f';
                    outBuffer.append(aChar);
                }
            } else outBuffer.append(aChar);
        }
        return outBuffer.toString();
    }
}