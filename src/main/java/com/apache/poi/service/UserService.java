package com.apache.poi.service;

import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.time.LocalDate;

public interface UserService {
    ByteArrayInputStream exportToExcel(LocalDate startDate, LocalDate endDate) throws IOException;
    void exportAndSendExcel(LocalDate startDate, LocalDate endDate, String email) throws Exception;
    void importUsersFromExcel(MultipartFile file) throws IOException;
}
