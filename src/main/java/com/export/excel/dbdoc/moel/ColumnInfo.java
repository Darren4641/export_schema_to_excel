package com.export.excel.dbdoc.moel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class ColumnInfo {
    private String tableName;
    private String columnName;
    private String columnType;
    private String isNullable;
    private String columnKey;
    private String extra;
    private String columnDefault;
    private String columnComment;


}