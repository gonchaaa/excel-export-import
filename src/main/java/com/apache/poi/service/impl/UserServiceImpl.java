package com.apache.poi.service.impl;

import com.apache.poi.entity.User;
import com.apache.poi.entity.UserExtraField;
import com.apache.poi.repository.UserRepository;
import com.apache.poi.service.EmailService;
import com.apache.poi.service.UserService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

@Service
@RequiredArgsConstructor
public class UserServiceImpl implements UserService {

    private final UserRepository userRepository;
    private final EmailService emailService;
    @Override
    public ByteArrayInputStream exportToExcel(LocalDate startDate, LocalDate endDate) throws IOException {

        List<User> userList = userRepository.findAllByInsertDateBetween(startDate, endDate);
        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        Sheet sheet = workbook.createSheet("Users");

        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("ID");
        headerRow.createCell(1).setCellValue("Name");
        headerRow.createCell(2).setCellValue("Email");
        headerRow.createCell(3).setCellValue("Insert Date");

        int rowNum = 1;

        for (User user:userList) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(user.getId());
            row.createCell(1).setCellValue(user.getUsername());
            row.createCell(2).setCellValue(user.getEmail());
            row.createCell(3).setCellValue(user.getInsertDate().toString());
        }
        workbook.write(outputStream);

        return new ByteArrayInputStream(outputStream.toByteArray());
    }

    @Override
    public void exportAndSendExcel(LocalDate startDate, LocalDate endDate, String email) throws Exception {
        ByteArrayInputStream excel = exportToExcel(startDate, endDate);
        emailService.sendEmailExcel(
                email,
                "Exported User Data",
                "Zəhmət olmasa əlavə edilmiş fayla baxın.",
                excel,
                "users.xlsx"
        );


    }
    @Override
    public void importUsersFromExcel(MultipartFile file) throws IOException {
        List<User> userList = new ArrayList<>();
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet sheet = workbook.getSheetAt(0);

        Row headerRow = sheet.getRow(0);
        int columnCount = headerRow.getLastCellNum();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                User user = new User();

                List<UserExtraField> extraFields = new ArrayList<>();

                for (int j = 0; j < columnCount; j++) {
                    String header = getCellValueAsString(headerRow.getCell(j));
                    String value = getCellValueAsString(row.getCell(j));

                    switch (header.toLowerCase()){
                        case "username":
                            user.setUsername(getCellValueAsString(row.getCell(2)));
                            break;
                        case "email":
                            user.setEmail(getCellValueAsString(row.getCell(0)));
                            break;
                        case "insert_date":
                            user.setInsertDate(LocalDate.parse(getCellValueAsString(row.getCell(1))));
                            break;
                        default:
                            UserExtraField extra = new UserExtraField();
                            extra.setFieldKey(header);
                            extra.setFieldValue(value);
                            extra.setUser(user);
                            extraFields.add(extra);
                    }

                }
                user.setExtraFields(extraFields);
                userList.add(user);
            }
        }

        userRepository.saveAll(userList);
    }
    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue().toLocalDate().toString();
                }
                return String.valueOf((long) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }


}
