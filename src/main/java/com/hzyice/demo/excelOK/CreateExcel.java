package com.hzyice.demo.excelOK;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Arrays;
import java.util.List;

/**
 * @Discreption 根据已有的Excel模板，修改模板内容生成新Excel
 */
@Slf4j
public class CreateExcel {


    public static void main(String[] args) throws IOException {
        //excle 2003
        //createXLS();
        //excle 2007
        //createXLSX();  // OK

        //projectCreateXLSX();
        /*int a = 1;
        int b = a;
        b += 2;
        System.out.println(a);
        System.out.println(b);*/

        //testSetSheestName();

        //log.info("水印模板生成完成...");


        String file = "c:abc\\e";
        String substring = file.substring(file.lastIndexOf("\\") + 1, file.length());
        System.out.println(substring.length());

    }



    /**
     *
     *(2003 xls后缀 导出)
     * @return void 返回类型
     * @author xsw
     * @2016-12-7上午10:44:00
     */
    public static void createXLS() throws IOException{
        //excel模板路径
        File fi=new File("D:\\offer_template.xls");
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(fi));
        //读取excel模板
        HSSFWorkbook wb = new HSSFWorkbook(fs);
        //读取了模板内所有sheet内容
        HSSFSheet sheet = wb.getSheetAt(0);

        //如果这行没有了，整个公式都不会有自动计算的效果的
        sheet.setForceFormulaRecalculation(true);


        //在相应的单元格进行赋值
        HSSFCell cell = sheet.getRow(11).getCell(6);//第11行 第6列
        cell.setCellValue(1);
        HSSFCell cell2 = sheet.getRow(11).getCell(7);
        cell2.setCellValue(2);
        sheet.getRow(12).getCell(6).setCellValue(12);
        sheet.getRow(12).getCell(7).setCellValue(12);
        //修改模板内容导出新模板
        FileOutputStream out = new FileOutputStream("D:/export.xls");
        wb.write(out);
        out.close();
    }
    /**
     *
     *(2007 xlsx后缀 导出)
     * @return void 返回类型
     * @author xsw
     * @2016-12-7上午10:44:30
     */
    public static void createXLSX() throws IOException{
        //excel模板路径
        File fi=new File("D:\\excel\\base\\style.xlsx");
        InputStream in = new FileInputStream(fi);
        //读取excel模板
        XSSFWorkbook wb = new XSSFWorkbook(in);
        //读取了模板内所有sheet内容
        XSSFSheet sheet = wb.getSheetAt(0);

        //如果这行没有了，整个公式都不会有自动计算的效果的
        sheet.setForceFormulaRecalculation(true);

        sheet.createRow(6).createCell(6).setCellValue(66);

        //修改模板内容导出新模板
        FileOutputStream out = new FileOutputStream("D:\\excel\\idea\\export.xlsx");
        wb.write(out);
        out.close();
    }



    // 按项目中
    public static void projectCreateXLSX() throws IOException{
        //excel模板路径
        File fi=new File("D:\\excel\\base\\style.xlsx");
        InputStream in = new FileInputStream(fi);
        XSSFWorkbook xssfSheets = new XSSFWorkbook(in);
        //读取excel模板
        //SXSSFWorkbook wb = new SXSSFWorkbook(xssfSheets);
        //读取了模板内所有sheet内容
        //Sheet sheet = wb.getSheetAt(0);
        //HSSFWorkbook   workbook   =   new HSSFWorkbook(in);
        //workbook.cloneSheet(1);
        List<String> strings = Arrays.asList("abc", "def", "ghi", "jkl", "mno");

        //wb.cloneSheet(1);




        //wb.c

        for (int i = 0; i < 5; i++) {
            xssfSheets.cloneSheet(0);
            xssfSheets.setSheetName(i, strings.get(i));

            log.info("i = ", i);
            //workbook.cloneSheet(0);
            //SXSSFSheet sheetAt = (SXSSFSheet)wb.createSheet(strings.get(i));
            //SXSSFSheet sheetAt = (SXSSFSheet)wb.getSheetAt(i);
            //sheetAt = (SXSSFSheet)wb.getSheetAt(0);
            //Sheet sheet = wb.cloneSheet(5);
            //wb.setSheetName(i, strings.get(i));
        }
        xssfSheets.removeSheetAt(5);

        //如果这行没有了，整个公式都不会有自动计算的效果的
        //sheet.setForceFormulaRecalculation(true);

        //sheet.createRow(6).createCell(6).setCellValue(66);

        SXSSFWorkbook wb = new SXSSFWorkbook(xssfSheets);



        //修改模板内容导出新模板
        FileOutputStream out = new FileOutputStream("D:\\excel\\idea\\export.xlsx");
        wb.write(out);
        //workbook.write(out);
        //xssfSheets.write(out);
        out.close();
    }


    public static void testSetSheestName() {

        //excel模板路径
        File fi=new File("D:\\excel\\base\\style.xlsx");

        XSSFWorkbook xssfSheets = null;

        try {
            InputStream in = new FileInputStream(fi);
            xssfSheets = new XSSFWorkbook(in);
        } catch (IOException e) {
            log.error("加载字节流异常...");
        }

        xssfSheets.setSheetName(0, "message");

        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(xssfSheets);

        SXSSFSheet sheet = (SXSSFSheet) sxssfWorkbook.getSheetAt(0);

        SXSSFCell cell = (SXSSFCell) sheet.createRow(0).createCell(0);

        cell.setCellType(SXSSFCell.CELL_TYPE_STRING);
        cell.setCellValue("错误了...");

        CellStyle messStyle = sxssfWorkbook.createCellStyle();

        cell.setCellStyle(messStyle);
        messStyle.setWrapText(true);
        Font font = sxssfWorkbook.createFont();
        messStyle.setFont(font);
        font.setColor(Font.COLOR_RED);
        sheet.autoSizeColumn(0);


        //修改模板内容导出新模板
        FileOutputStream out = null;
        try {
            out = new FileOutputStream("D:\\excel\\idea\\export.xlsx");
            sxssfWorkbook.write(out);
            //workbook.write(out);
            //xssfSheets.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }


    }


    public static void testExcelName() {

        File tempFile2 = new File("D:\\excel\\base\\style.xlsx");;

        XSSFWorkbook xssfSheets = null;

        try {
            InputStream in = new FileInputStream(tempFile2);
            xssfSheets = new XSSFWorkbook(in);
        } catch (IOException e) {
            log.error("加载字节流异常...");
        }

        xssfSheets.setSheetName(0, "message");

        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(xssfSheets);

        //写入临时xlsx
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(tempFile2);
            sxssfWorkbook.write(fileOutputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }




    }







}
