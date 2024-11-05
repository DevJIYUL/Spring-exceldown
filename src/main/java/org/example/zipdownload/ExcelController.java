package org.example.zipdownload;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpRequest;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import java.io.ByteArrayOutputStream;

@RestController
public class ExcelController {
    @GetMapping("exceldown")
    public void getExcel(HttpServletRequest request, HttpServletResponse response) throws Exception{
        List<Account> accounts = new ArrayList<>();
        for (int i = 0; i < 3; i++) {
            accounts.add(Account.builder().acctno("10-112-1264"+i).name("pay-"+i).acctholder("khan-"+i).build());
        }

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("첫번째 시트");
        Row row = null;
        Cell cell = null;
        int rowNum = 0;

        // Header
        row = sheet.createRow(rowNum++);
        cell = row.createCell(0);
        cell.setCellValue("계좌번호");
        cell = row.createCell(1);
        cell.setCellValue("예금주");
        cell = row.createCell(2);
        cell.setCellValue("이름");

        // Body
        for (int i=0; i<3; i++) {
            row = sheet.createRow(rowNum++);
            cell = row.createCell(0);
            cell.setCellValue(accounts.get(i).getAcctno());
            cell = row.createCell(1);
            cell.setCellValue(accounts.get(i).getAcctholder());
            cell = row.createCell(2);
            cell.setCellValue(accounts.get(i).getName());
        }

        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=studentList.xls");

        wb.write(response.getOutputStream());
        wb.close();
    }
}
