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
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@RestController
public class ZipFileExcelController {
    @GetMapping("zipfileexceldown")
    public void getExcel(HttpServletRequest request, HttpServletResponse response) throws Exception{
        int rowCount = 340;
        int pivot = 100;

        Workbook wb =null;
        Sheet sheet;
        File currDir;
        currDir = new File(".");                // 현재 프로젝트 경로를 가져옴
        String path = currDir.getAbsolutePath();

        List<String > fileNames = new ArrayList<>();

        for (int i = 0; i < (rowCount/pivot)+1; i++) {
            List<Account> accounts = new ArrayList<>();
            for (int j = 0; j < pivot; j++) {
                if(i*pivot+j>rowCount)break;
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
            for (int q=0; q<accounts.size(); q++) {
                row = sheet.createRow(rowNum++);
                cell = row.createCell(0);
                cell.setCellValue(accounts.get(q).getAcctno());
                cell = row.createCell(1);
                cell.setCellValue(accounts.get(q).getAcctholder());
                cell = row.createCell(2);
                cell.setCellValue(accounts.get(q).getName());
            }
            String fileLocation = path.substring(0, path.length() - 1) + "temp"+(i+1)+".xlsx";    // 파일명 설정

            fileNames.add("temp"+(i+1)+".xlsx");

            try (FileOutputStream fileout = new FileOutputStream(fileLocation)){
                wb.write(fileout);
            }catch (Exception e){

            }finally {
                wb.close();
            }
        }

        response.setContentType("application/zip");
        response.setHeader("Content-Disposition", "attachment; filename=\"alphabet.zip\""); // 변경된 부분


//        List<File> goZip = new ArrayList<>();
//        for (String file : fileNames){
//            File excel = new File(path,file);
//            goZip.add(excel);
//        }

        File zipFile = new File(path,"temp.zip" );
        byte[] buf = new byte[4096];
//        try (ZipOutputStream zipOut = new ZipOutputStream(new FileOutputStream(zipFile))) {  // 집 파일 다운로드
        try (ZipOutputStream zipOut = new ZipOutputStream(response.getOutputStream())) {  // 스트리밍 방식으로 ZIP 파일을 직접 반환
            for (String fileName : fileNames) {
                File excelFile = new File(path, fileName);
                try (FileInputStream in = new FileInputStream(excelFile)) {
                    ZipEntry ze = new ZipEntry(excelFile.getName());
                    zipOut.putNextEntry(ze);

                    int len;
                    while ((len = in.read(buf)) > 0) {
                        zipOut.write(buf, 0, len);
                    }
                    zipOut.closeEntry();
                }
                // 엑셀 파일 삭제하는 로직
//                if(excelFile.exists()){
//                    if(excelFile.delete()){
//                        System.out.println(excelFile.getName()+" 파일 삭제되었음");
//                    }else{
//                        System.out.println(excelFile.getName()+" 파일 삭제되지 않음");
//                    }
//                }
            }
//            zipOut.finish();
        }

    }
}
