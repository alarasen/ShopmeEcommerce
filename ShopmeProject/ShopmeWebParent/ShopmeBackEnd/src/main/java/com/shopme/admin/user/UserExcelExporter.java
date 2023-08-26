package com.shopme.admin.user;

import com.shopme.common.entity.User;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

public class UserExcelExporter extends AbstractExporter {

    public void export(List<User> users, HttpServletResponse response) throws IOException {
        String[] excelHeader = {"User ID", "E-mail", "First Name", "Last Name", "Roles", "Enabled"};

        String[] fieldMapping = {"id", "email", "firstName", "lastName", "roles", "enabled"};

        // HTTP yanıt başlığını ayarla
        super.setResponseHeader(response, "application/octet-stream", ".xlsx");

        // Yeni bir Excel işbook (çalışma kitabı) oluştur
        XSSFWorkbook workbook = new XSSFWorkbook();

        // "User List" adında bir sayfa (sheet) oluştur
        Sheet sheet = workbook.createSheet("User List");

        // Başlık (header) satırını oluştur
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < excelHeader.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(excelHeader[i]);
        }

        for (int i = 0; i < users.size(); i++) {
            Row data = sheet.createRow(i + 1); // Başlık satırı dışında veri satırları başlangıcı
            User user = users.get(i);

            for (int j = 0; j < fieldMapping.length; j++) {
                Cell cell = data.createCell(j);
                String fieldName = fieldMapping[j];

                // "User" nesnesinden ilgili alanı al
                String value = getUserFieldValue(user, fieldName);

                // Hücreye değeri yerleştir
                cell.setCellValue(value);
            }
        }

        try {
            // Excel dosyasını yanıt olarak gönder
            ServletOutputStream outputStream = response.getOutputStream();
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private String getUserFieldValue(User user, String fieldName) {
        if (fieldName.equals("id"))
            return String.valueOf(user.getId());
        if (fieldName.equals("email"))
            return String.valueOf(user.getEmail());
        if (fieldName.equals("firstName"))
            return String.valueOf(user.getFirstName());
        if (fieldName.equals("lastName"))
            return String.valueOf(user.getLastName());
        if (fieldName.equals("roles"))
            return String.valueOf(user.getRoles());
        if (fieldName.equals("enabled"))
            return String.valueOf(user.getEnabled());

        return "something is wrong!";
    }


}

