package org.example.constructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.NonNull;
import lombok.Setter;

@NoArgsConstructor
@Getter
@Setter
public class Change_Message {
    @NonNull
    private String variable_name;

    private String eng;

    private String eng2;

    private String zh_ch;

    private String zh_hk;

    private String new_variable_name;

    private String new_eng;

    private String new_eng2;

    private String new_zh_ch;

    private String new_zh_hk;

}