package com.hzyice.demo.excelOK;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import javax.servlet.http.HttpServletResponse;
import java.io.*;

/* web 测试水印模块里写数据，生成新的excel*/

@Controller
@Slf4j
public class ExportController {
    @RequestMapping(value = "/exportSheet")
    public void exportExcel(HttpServletResponse response) {

        //excel模板路径
        File fi = new File("/home/template/style.xlsx");
        InputStream in = null;
        try {
            in = new FileInputStream(fi);
            //读取excel模板
            XSSFWorkbook wb = new XSSFWorkbook(in);
            //读取了模板内所有sheet内容
            XSSFSheet sheet = wb.getSheetAt(0);

            //如果这行没有了，整个公式都不会有自动计算的效果的
            sheet.setForceFormulaRecalculation(true);

            sheet.createRow(6).createCell(6).setCellValue(66);

            //输出Excel文件
            OutputStream output = response.getOutputStream();
            response.reset();
            response.setHeader("Content-disposition", "attachment; filename=template.xls");
            response.setContentType("application/msexcel");
            wb.write(output);
            output.close();
        } catch (Exception e) {
            log.info("输出excel文件失败...");
            e.printStackTrace();
        }
    }


}

