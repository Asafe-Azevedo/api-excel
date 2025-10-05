package org.example;


import org.apache.poi.ss.usermodel.*;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/excel")
public class ExcelController {

    @PostMapping("/upload")
    public ResponseEntity<?> uploadExcel(@RequestParam("file") MultipartFile file){
        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = WorkbookFactory.create(inputStream)){

            Sheet sheet = workbook.getSheetAt(0);
            List<Map<String, Object>> dataList = new ArrayList<>();

            Row headerRow = sheet.getRow(0);
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow){
                headers.add(cell.getStringCellValue());
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++){
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Map<String, Object> data = new LinkedHashMap<>();
                for (int j = 0; j < headers.size(); j++){
                    Cell cell = row.getCell(j);
                    Object value = getCellValue(cell);
                    data.put(headers.get(j), value);
                }
                dataList.add(data);
            }
            return ResponseEntity.ok(dataList);
        } catch (Exception e){
            e.printStackTrace();
            return ResponseEntity.badRequest().body("Erro ao processar arquivo: " + e.getMessage());
        }
    }

    private Object getCellValue(Cell cell){
        if (cell == null) return null;
        return switch (cell.getCellType()){
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> cell.getNumericCellValue();
            case BOOLEAN -> cell.getBooleanCellValue();
            default -> null;
        };
    }
}
