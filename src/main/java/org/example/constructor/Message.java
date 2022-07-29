package org.example.constructor;
import lombok.*;

@AllArgsConstructor
@NoArgsConstructor
@Getter
@Setter
public class Message {
    @NonNull
    private String variable_name;
    private String eng;
    private String eng2;
    private String zh_ch;
    private String zh_hk;
}
