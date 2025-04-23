package com.apache.poi.entity;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Entity
@Data
@NoArgsConstructor
@AllArgsConstructor
@Table(name = "user_extra_fields")
public class UserExtraField {

    @Id
    @GeneratedValue
    private Long id;

    private String fieldKey;
    private String fieldValue;

    @ManyToOne
    @JoinColumn(name = "user_id")
    private User user;
}
