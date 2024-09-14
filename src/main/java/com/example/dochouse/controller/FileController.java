package com.example.dochouse.controller;

import com.example.dochouse.service.FileConversionService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@RequestMapping("/api/files")
//@CrossOrigin(origins = "chrome-extension://lldgdmkecjecbahnmjaapeeeecooajdp")
@CrossOrigin(origins = "*")
public class FileController {

    @Autowired
    private FileConversionService conversionService;

    @PostMapping("/convert-to-pdf")
    public ResponseEntity<?> convertToPdf(@RequestParam("file") MultipartFile file,
                                          @RequestParam("type") String type) {
        try {
            byte[] pdfBytes = new byte[0];
            switch (type.toLowerCase()) {
                case "word":
                    pdfBytes = conversionService.convertWordToPdf(file);
                    break;
                case "ppt":
                    pdfBytes = conversionService.convertPptToPdf(file);
                    break;
                case "txt":
                    pdfBytes = conversionService.convertTextToPdf(file);
                    break;
                case "excel":
                    pdfBytes = conversionService.convertExcelToPdf(file);
                    break;
                case "image":
                    pdfBytes = conversionService.convertImageToPdf(file);
                    break;
                default:
                    return ResponseEntity.badRequest().body("Unsupported file type.");
            }

            return ResponseEntity.ok()
                    .contentType(MediaType.APPLICATION_PDF)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"converted.pdf\"")
                    .body(pdfBytes);
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(e.getMessage());
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }


//CONVERTS JPG TO JPEG, JPEG TO JPG, JPG/JPEG TO PNG, BUT PNG TO JPG/JPEG NOT HAPPENING.
    @PostMapping("/convert-image-format")
    public ResponseEntity<?> convertImageFormat(@RequestParam("file") MultipartFile file,
                                                @RequestParam("format") String format) {
        try {
            byte[] convertedBytes = conversionService.convertImageFormat(file, format);
            return ResponseEntity.ok()
                    .contentType(MediaType.IMAGE_JPEG)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"converted." + format + "\"")
                    .body(convertedBytes);
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(e.getMessage());
        }
    }
}
