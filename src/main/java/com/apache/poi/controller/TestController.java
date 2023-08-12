package com.apache.poi.controller;

import com.apache.poi.service.GenerateExcelService;
import lombok.AllArgsConstructor;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
@RequestMapping("/test")
@AllArgsConstructor
public class TestController {
    private final GenerateExcelService generateExcelService;

    @GetMapping("/home")
    public String getHome(){
        return "home";
    }

    @GetMapping("/download")
    public ResponseEntity<Resource> getFile() {
        String filename = "Merge_cell_handle.xlsx";
        InputStreamResource file = new InputStreamResource(generateExcelService.load());
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
                .body(file);
    }
}
