package com.hzyice.demo.excelOK;

import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
@Slf4j
public class testExcedl {

    /*给excel加水印---图片的方式*/
    public static void main(String[] args) {
        ExcelShuiyinUtil util = new ExcelShuiyinUtil();
        util.OpenExcel("D:\\excel\\style.xlsx", false);
        util.setBlackGroudPrituce("D:\\excel\\sy.png");
        util.CloseExcel(true);

        log.info("水印添加成功...");
    }

    @RequestMapping("/sy")
    @ResponseBody
    public String exportExcel(){
        try {
            ExcelShuiyinUtil util = new ExcelShuiyinUtil();
            util.OpenExcel("/home/template/sy.xls", false);
            util.setBlackGroudPrituce("/home/template/sy.png");
            util.CloseExcel(true);
            return "水印添加成功...";
        } catch (Exception e) {
            return "水印添加失败...";
        }
    }
}
