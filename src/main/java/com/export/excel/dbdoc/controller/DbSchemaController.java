package com.export.excel.dbdoc.controller;

import com.export.excel.dbdoc.service.DbSchemaService;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@RestController
@RequiredArgsConstructor
public class DbSchemaController {

    private final DbSchemaService dbSchemaService;

    @GetMapping("/api/db-schema/excel")
    public ResponseEntity<ByteArrayResource> downloadDbSchemaExcel() throws IOException {
        ByteArrayOutputStream excelData = dbSchemaService.generateExcel();
        ByteArrayResource resource = new ByteArrayResource(excelData.toByteArray());

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"DB_Schema.xlsx\"")
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .contentLength(resource.contentLength())
                .body(resource);
    }
}
