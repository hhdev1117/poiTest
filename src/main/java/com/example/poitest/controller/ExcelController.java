package com.example.poitest.controller;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;


import java.io.*;

@Controller
public class ExcelController {

    @RequestMapping("download")
    public void create(HttpServletRequest request, HttpServletResponse response)
            throws Exception {

        // 1. Excel Workbook(파일)을 생성한다. (Template을 읽어온다)
        SXSSFWorkbook workbook = getSxssfWorkbook(request, "sampleTemplate." +
                "xlsx");

        // 2. Template에는 이미 Sheet가 존재하므로, Sheet Index로 데이터를 적용할 Sheet를 가져온다.
        SXSSFSheet sheet = workbook.getSheetAt(0);

        // 3. 생성된 Sheet에 Row를 만든다. 해당 Template에는 0번 Row가 있으므로 1번 Row부터 작성한다.
        SXSSFRow row = sheet.createRow(1);

        // 4. 생성된 Row에 cell을 만든다.
        SXSSFCell cell = row.createCell(0);

        // 5. 생성된 Cell에 값을 입력한다.
        cell.setCellValue("value");

        runToExcel(workbook, response, "123");
        //download(workbook, response);
    }

    public SXSSFWorkbook getSxssfWorkbook(HttpServletRequest request, String templateName) throws IOException {
        // 1. 템플릿을 받을 InputStream과, 반환해줄 SXSSFWorkbook을 생성한다.
        InputStream inputStream = null;
        SXSSFWorkbook sxssfWorkbook = null;
        try {
            // 2. template file의 경로를 찾아서 inputStream으로 받는다.
            String templatePath = request.getSession().getServletContext().getRealPath("/WEB-INF/template/");
            inputStream = new BufferedInputStream(new FileInputStream(templatePath + templateName));

            // 3. 해당 InputStream(읽어온 Template)으로 XSSFWorkBook을 생성한다.
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(inputStream);

            // 4. 생성한 XSSFWorkBook을 SXSSFWorkbook으로 변환한다.
            sxssfWorkbook = new SXSSFWorkbook(xssfWorkbook, 10);
        } catch(Exception e) {
            // Exception 처리하세요.
        } finally {
            // 5. 사용한 Stream을 닫는다.
            //if(inputStream != null) inputStream.close();
        }

        return sxssfWorkbook;
    }


    public void download(SXSSFWorkbook workbook, HttpServletResponse response) throws IOException {
        // 1. OutputStream 생성
        OutputStream outputStream = null;
        try {
            // 2. mime-type을 설정한다.
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            // 3. 파일명을 설정한다.
            response.setHeader("Content-Disposition", "Attachment; Filename=sample.xlsx");

            // 4. response의 OutputStream에 생성한 workbook을 write한다.
            outputStream = new BufferedOutputStream(response.getOutputStream());
            workbook.write(outputStream);
            outputStream.flush();

        } catch (Exception e) {
            // Exception 처리하세요.
        } finally {
            // 5. 사용한 Stream을 닫는다.
            if(outputStream != null) outputStream.close();
        }
    }

    public void encodeDownload(SXSSFWorkbook workbook, HttpServletResponse response, String password) throws IOException {
        ByteArrayOutputStream byteArrayOutputStream = null;
        InputStream inputStream = null;
        POIFSFileSystem poifsFileSystem = null;
        OPCPackage opcPackage = null;

        try {
            // 2. mime-type을 설정한다.
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition", "Attachment; Filename=sample.xlsx");
            //response.setHeader("Set-Cookie", "fileDownload=true; path=/");

            // 1. SXSSFWorkbook을 ByteArrayOutputStream으로 내보낸다.
            byteArrayOutputStream = new ByteArrayOutputStream();
            workbook.write(byteArrayOutputStream);

            // 2. ByteArrayOutputStream을 InputStream으로 가져온다.
            inputStream = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());

            poifsFileSystem = new POIFSFileSystem();

            opcPackage = OPCPackage.open(new ByteArrayInputStream(byteArrayOutputStream.toByteArray()));

            Encryptor encryptor = new EncryptionInfo(EncryptionMode.agile).getEncryptor();
            encryptor.confirmPassword(password);

            opcPackage.save(encryptor.getDataStream(poifsFileSystem));

            OutputStream tempstream = response.getOutputStream();
            poifsFileSystem.writeFilesystem(tempstream);

        } catch (Exception e) {
            // Exception 처리하세요.
        } finally {
            if(byteArrayOutputStream != null) byteArrayOutputStream.close();
            if(inputStream != null) inputStream.close();
            if(poifsFileSystem != null) poifsFileSystem.close();
            if(opcPackage != null) opcPackage.close();
        }
    }

    public void runToExcel(SXSSFWorkbook workbook, HttpServletResponse response, String excelpass) throws Exception {
        try {
            // 엑셀다운 response 설정
            response.setContentType("Application/Msexcel");
            response.setHeader("Content-Disposition",
                    "ATTachment; Filename=sample.xlsx");

            // sxssf는 읽기가 불가능하므로 워크북을 읽어서 암호를 입히는것이 안된다.
            // 따라서 inputstream에 저장후 다시 암호를 입힐 xssf를 생성한다.
            ByteArrayOutputStream tempOs = new ByteArrayOutputStream();
            workbook.write(tempOs);
            InputStream in = new ByteArrayInputStream(tempOs.toByteArray());

            // 파일을 Output하기위한 객체
            POIFSFileSystem fs = new POIFSFileSystem();
            // 파일 암호화
            OPCPackage opc = OPCPackage.open(in);
            OutputStream os = getEncryptOutputStream(fs, excelpass);
            opc.save(os);
            opc.close();

            response.setHeader("Set-Cookie", "fileDownload=true; path=/");
            OutputStream resOS = response.getOutputStream();
            fs.writeFilesystem(resOS);

            // 썻던 stream들 다 닫아줍니다.
            tempOs.close();
            in.close();
            resOS.close();
            fs.close();
            workbook.close();

        }
        catch (NullPointerException e) {
        }
        catch (IOException e) {
        }
        catch (Exception e) {
        }
    }

    private OutputStream getEncryptOutputStream(POIFSFileSystem fileSystem, String excelpass) {
        try {
            Encryptor enc = new EncryptionInfo(EncryptionMode.agile).getEncryptor();
            enc.confirmPassword(excelpass);
            return enc.getDataStream(fileSystem);
        }
        catch (NullPointerException e) {
        }
        catch (IOException e) {
        }
        catch (Exception e) {
        }
        return null;
    }
}
