package com.apache.poi.service;

import jakarta.mail.MessagingException;

import java.io.ByteArrayInputStream;

public interface EmailService {
    void sendEmailExcel(String to, String subject, String text, ByteArrayInputStream attachment, String filename) throws MessagingException;

}
