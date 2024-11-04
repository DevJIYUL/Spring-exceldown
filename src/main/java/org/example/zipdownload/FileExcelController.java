package org.example.zipdownload;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

@RestController
public class FileExcelController {
    @GetMapping("fileexceldown")
    public void getExcel(HttpServletRequest request, HttpServletResponse response) throws Exception{
        int rowCount = 340;
        int pivot = 100;

        Workbook wb =null;
        Sheet sheet;
        File currDir;
        List<String> fileDir = new ArrayList<>();
        for (int i = 0; i < (rowCount/pivot)+1; i++) {
            List<Account> accounts = new ArrayList<>();
            for (int j = 0; j < pivot; j++) {
                if(i*pivot+j> rowCount)break;
                accounts.add(Account.builder().acctno(String.valueOf(i*pivot+j)).name("페이"+i).acctholder("kain"+i).build());
            }

            System.out.println(i+"번째");
            System.out.println(accounts);
            wb = new XSSFWorkbook();
            sheet = wb.createSheet("첫번째 시트");
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
            for (int q=0; q<pivot; q++) {
                row = sheet.createRow(rowNum++);
                cell = row.createCell(0);
                cell.setCellValue(accounts.get(q).getAcctno());
                cell = row.createCell(1);
                cell.setCellValue(accounts.get(q).getAcctholder());
                cell = row.createCell(2);
                cell.setCellValue(accounts.get(q).getName());
            }
            currDir = new File(".");                // 현재 프로젝트 경로를 가져옴
            String path = currDir.getAbsolutePath();
            String fileLocation = path.substring(0, path.length() - 1) + "temp"+(i+1)+".xls";    // 파일명 설정

            fileDir.add(fileLocation);

            try (FileOutputStream fileout = new FileOutputStream(fileLocation)){
                wb.write(fileout);
            }catch (Exception e){

            }finally {
                wb.close();
            }
        }
//        response.setContentType("ms-vnd/excel");
//        response.setHeader("Content-Disposition", "attachment;filename=studentList.xls");


//        FileOutputStream fileOutputStream = null;
//        for (String fileName : fileDir){
//             fileOutputStream = new FileOutputStream(fileName);        // 파일 생성
//        }
//        wb.write(fileOutputStream);                                            // 엑셀파일로 작성
//        wb.close();
//
//        System.out.println(fileDir);


    }
}
