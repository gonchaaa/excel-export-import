package com.apache.poi.controller;

import com.apache.poi.service.UserService;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.time.LocalDate;

@RestController
@RequiredArgsConstructor
@RequestMapping("/users")
public class UserController {

    private final UserService userService;

    @GetMapping
    public ResponseEntity<InputStreamResource> exportToExcel(@RequestParam LocalDate startDate,
                                                             @RequestParam LocalDate endDate) throws IOException {

        ByteArrayInputStream inputStream = userService.exportToExcel(startDate, endDate);

        HttpHeaders headers = new HttpHeaders();
        headers.add("Content-Disposition", "attachment; filename=users.xlsx");

        return ResponseEntity.ok().headers(headers)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(new InputStreamResource(inputStream));
    }

    @GetMapping("/send-email")
    public ResponseEntity<String> exportAndSendByEmail(
            @RequestParam LocalDate startDate,
            @RequestParam LocalDate endDate,
            @RequestParam String email
    ) throws Exception {
        try {
            userService.exportAndSendExcel(startDate, endDate, email);
            return ResponseEntity.ok("Email gonderildi");
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("Failed to send email: " + e.getMessage());
        }
    }

    @PostMapping()
    public ResponseEntity<String> importUsersFromExcel(@RequestParam("file") MultipartFile file) {
        try {
            userService.importUsersFromExcel(file);
            return ResponseEntity.ok("Data uğurla import edildi.");
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("Import zamanı xəta baş verdi: " + e.getMessage());
        }
    }


}
