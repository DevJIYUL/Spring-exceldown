package org.example.zipdownload;

import lombok.*;

@Builder
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@ToString
public class Account {
    private String acctno;
    private String acctholder;
    private String name;
}
