package com.apache.poi.service;

import com.apache.poi.util.ExcelUtil;
import lombok.AllArgsConstructor;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;

@Service
@AllArgsConstructor
public class GenerateExcelService {
    private final ExcelUtil excelUtil;

    public ByteArrayInputStream load() {
        ByteArrayInputStream inputStream = excelUtil.tutorialsToExcel();
        return inputStream;
    }
}
