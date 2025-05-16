package com.export.excel.dbdoc.service;

import com.export.excel.dbdoc.moel.ColumnInfo;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
public class DbSchemaService {

    private final JdbcTemplate jdbcTemplate;
    @Value("${spring.datasource.url}")
    private String dataSourceUrl;


    /**
     * DB 메타데이터를 조회하여
     * "테이블 명 영문 / 테이블 명 한글"을 한 줄에 배치하고
     * 그 뒤 바로 헤더-컬럼 정보를 출력하며,
     * 테이블 간 공백 한 줄을 두는 Excel 파일을 생성
     */
    public ByteArrayOutputStream generateExcel() throws IOException {
        // 1. DB 스키마명 추출
        String schemaName = extractSchemaFromUrl(dataSourceUrl);

        // 2. information_schema.columns 에서 컬럼 메타데이터 조회
        String sql = """
                SELECT
                       c.TABLE_NAME as TABLE_NAME,
                       t.TABLE_COMMENT as TABLE_COMMENT,
                       c.COLUMN_NAME as COLUMN_NAME,
                       c.COLUMN_TYPE as COLUMN_TYPE,
                       c.IS_NULLABLE as IS_NULLABLE,
                       c.COLUMN_KEY as COLUMN_KEY,
                       c.EXTRA as EXTRA,
                       c.COLUMN_DEFAULT as COLUMN_DEFAULT,
                       c.COLUMN_COMMENT as COLUMN_COMMENT
                FROM information_schema.columns c
                JOIN information_schema.tables t
                ON c.TABLE_SCHEMA = t.TABLE_SCHEMA
                AND c.TABLE_NAME = t.TABLE_NAME
                WHERE c.table_schema = ?
                ORDER BY TABLE_NAME, ORDINAL_POSITION
                """;

        List<ColumnInfo> columnList = jdbcTemplate.query(sql,
                new Object[]{ schemaName },
                (rs, rowNum) -> {
                    ColumnInfo info = new ColumnInfo();
                    info.setTableName(rs.getString("TABLE_NAME"));
                    info.setTableComment(rs.getString("TABLE_COMMENT"));
                    info.setColumnName(rs.getString("COLUMN_NAME"));
                    info.setColumnType(rs.getString("COLUMN_TYPE"));
                    info.setIsNullable(rs.getString("IS_NULLABLE"));
                    info.setColumnKey(rs.getString("COLUMN_KEY"));
                    info.setExtra(rs.getString("EXTRA"));
                    info.setColumnDefault(rs.getString("COLUMN_DEFAULT"));
                    info.setColumnComment(rs.getString("COLUMN_COMMENT"));
                    return info;
                }
        );

        // 3. 테이블별 그룹화 (순서 유지)
        Map<String, List<ColumnInfo>> tableMap = columnList.stream()
                .collect(Collectors.groupingBy(ColumnInfo::getTableName, LinkedHashMap::new, Collectors.toList()));

        // 4. Workbook/Sheet 생성
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("DB_Schema");

        // 각 열별로 너비 설정 (단위: 1/256th of a character width)
        sheet.setColumnWidth(0, 3000); // No
        sheet.setColumnWidth(1, 5000); // 컬럼명
        sheet.setColumnWidth(2, 4000); // 속성명
        sheet.setColumnWidth(3, 5000); // 데이터 타입(길이)
        sheet.setColumnWidth(4, 4000); // NULL 허용
        sheet.setColumnWidth(5, 4000); // 자동 증가
        sheet.setColumnWidth(6, 4000); // KEY
        sheet.setColumnWidth(7, 4000); // 기본값
        sheet.setColumnWidth(8, 6000); // Comment


        // 스타일 준비
        CellStyle labelStyle = createLabelStyle(workbook);    // 레이블(회색)
        CellStyle valueStyle = createValueStyle(workbook);    // 값(흰색)
        CellStyle headerStyle = createHeaderStyle(workbook);  // 헤더(회색)
        CellStyle bodyStyle = createBodyStyle(workbook);      // 본문(흰색)
        CellStyle bodyCenterStyle = createBodyCenterStyle(workbook); // 본문(가운데 정렬, 흰색)

        int rowIndex = 0;

        // 5. 테이블별로 반복
        for (Map.Entry<String, List<ColumnInfo>> entry : tableMap.entrySet()) {
            String tableName = entry.getKey();
            List<ColumnInfo> cols = entry.getValue();

            // (A) Row 0 : 테이블 명 (영문, 한글) 레이아웃
            Row row0 = sheet.createRow(rowIndex++);
            row0.setHeightInPoints(15);

            // 1) (0..1) 병합: "테이블 명 영문" (레이블)
            Cell cellLabelEng = row0.createCell(0);
            cellLabelEng.setCellStyle(labelStyle);
            cellLabelEng.setCellValue("테이블 명 영문");
            // 병합 영역
            sheet.addMergedRegion(new CellRangeAddress(
                    row0.getRowNum(), row0.getRowNum(), 0, 1
            ));
            // 병합된 나머지 셀(1)은 비워둬도 됨 (스타일 적용 시 편의상 같이 설정 가능)
            Cell cellLabelEngDummy = row0.createCell(1);
            cellLabelEngDummy.setCellStyle(labelStyle);

            // 2) (2..4) 병합: 테이블 명 영문 값 (흰색)
            Cell cellValueEng = row0.createCell(2);
            cellValueEng.setCellStyle(valueStyle);
            cellValueEng.setCellValue(tableName);
            sheet.addMergedRegion(new CellRangeAddress(
                    row0.getRowNum(), row0.getRowNum(), 2, 4
            ));
            Cell cellValueEng2 = row0.createCell(3);
            cellValueEng2.setCellStyle(valueStyle);
            Cell cellValueEng3 = row0.createCell(4);
            cellValueEng3.setCellStyle(valueStyle);

            // 3) (5..7) 병합 : "테이블 명 한글" (레이블, 한 칸)
            Cell cellLabelKor = row0.createCell(5);
            cellLabelKor.setCellStyle(labelStyle);
            cellLabelKor.setCellValue("테이블 명 한글");
            sheet.addMergedRegion(new CellRangeAddress(
                    row0.getRowNum(), row0.getRowNum(), 5, 7
            ));
            Cell cellValueKor2 = row0.createCell(6);
            cellValueKor2.setCellStyle(valueStyle);
            Cell cellValueKor3 = row0.createCell(7);
            cellValueKor3.setCellStyle(valueStyle);

            // 4) (8) : 테이블 명 한글 값 (흰색)
            Cell cellValueKor = row0.createCell(8);
            cellValueKor.setCellStyle(valueStyle);
            cellValueKor.setCellValue(cols.get(0).getTableComment()); // 실제 한글 테이블명 필요하면 여기 set

            // (B) 바로 아래줄 Row 1 : 헤더 (회색)
            Row headerRow = sheet.createRow(rowIndex++);
            headerRow.setHeightInPoints(15);
            String[] headers = {
                    "No", "컬럼명", "속성명", "데이터 타입",
                    "NULL 허용", "자동증가", "KEY", "기본값", "Comment"
            };
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle); // 요청대로 동일 회색
            }

            // (C) 컬럼 데이터 (흰색)
            int no = 1;
            for (ColumnInfo col : cols) {
                Row row = sheet.createRow(rowIndex++);
                row.setHeightInPoints(15);  // 데이터 행 높이 조정
                int colIdx = 0;

                // No
                Cell cNo = row.createCell(colIdx++);
                cNo.setCellValue(no++);
                cNo.setCellStyle(bodyCenterStyle);

                // 컬럼명
                Cell cName = row.createCell(colIdx++);
                cName.setCellValue(col.getColumnName());
                cName.setCellStyle(bodyStyle);

                // 속성명 (공란)
                Cell cAttr = row.createCell(colIdx++);
                cAttr.setCellValue("");
                cAttr.setCellStyle(bodyStyle);

                // 데이터타입
                Cell cType = row.createCell(colIdx++);
                cType.setCellValue(col.getColumnType());
                cType.setCellStyle(bodyCenterStyle);

                // NULL 허용
                Cell cNull = row.createCell(colIdx++);
                cNull.setCellValue(col.getIsNullable());
                cNull.setCellStyle(bodyCenterStyle);

                // 자동증가 여부
                String autoInc = (col.getExtra() != null && col.getExtra().contains("auto_increment")) ? "YES" : "";
                Cell cAuto = row.createCell(colIdx++);
                cAuto.setCellValue(autoInc);
                cAuto.setCellStyle(bodyCenterStyle);

                // KEY
                Cell cKey = row.createCell(colIdx++);
                cKey.setCellValue(col.getColumnKey());
                cKey.setCellStyle(bodyCenterStyle);

                // 기본값
                Cell cDef = row.createCell(colIdx++);
                cDef.setCellValue(col.getColumnDefault() == null ? "" : col.getColumnDefault());
                cDef.setCellStyle(bodyCenterStyle);

                // Comment
                Cell cComment = row.createCell(colIdx++);
                cComment.setCellValue(col.getColumnComment() == null ? "" : col.getColumnComment());
                cComment.setCellStyle(bodyStyle);
            }

            // (D) 테이블 끝나면 빈 행 1줄 → 다음 테이블과 간격
            rowIndex++;
        }

        // 6. Workbook → ByteArrayOutputStream
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        workbook.write(bos);
        workbook.close();
        return bos;
    }

    // ---------------------- 스타일 설정 메서드들 ----------------------

    /**
     * 레이블(테이블 명 영문/한글) 스타일: 회색 배경, Bold, 가운데 정렬, 테두리
     */
    private CellStyle createLabelStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();

        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    /**
     * 값(테이블 명 실제 값) 스타일: 흰색 배경, 테두리, 좌측 정렬
     */
    private CellStyle createValueStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();

        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    /**
     * 헤더(No, 컬럼명, 속성명, ...) 스타일: 요청대로 회색 배경, Bold, 중앙 정렬, 테두리
     */
    private CellStyle createHeaderStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();

        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    /**
     * 본문(컬럼 데이터) 스타일: 흰색 배경, 테두리, 왼쪽 정렬
     */
    private CellStyle createBodyStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();

        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    /**
     * 본문(컬럼 데이터, 가운데 정렬) 스타일: 흰색 배경, 중앙 정렬, 테두리
     */
    private CellStyle createBodyCenterStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    /**
     * spring.datasource.url 에서 DB 스키마명 추출
     * 예) "jdbc:mariadb://localhost:3306/my_db?useUnicode=true" → "my_db"
     */
    private String extractSchemaFromUrl(String url) {
        int qmIndex = url.indexOf("?");
        if (qmIndex != -1) {
            url = url.substring(0, qmIndex);
        }
        int slashIndex = url.lastIndexOf("/");
        if (slashIndex != -1 && slashIndex < url.length() - 1) {
            return url.substring(slashIndex + 1);
        }
        throw new IllegalArgumentException("Invalid DataSource URL: " + url);
    }
}
