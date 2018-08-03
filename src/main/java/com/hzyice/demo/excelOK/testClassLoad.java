package com.hzyice.demo.excelOK;

import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

@Slf4j
public class testClassLoad {

    private static InputStream exceInputStream;

    // 加载配置文件读取字节流
    static{
        Properties properties = new Properties();

        InputStream excelStream = testClassLoad.class.getClassLoader().getResourceAsStream("excel-template.properties");
        try {
            properties.load(excelStream);
            String filePath = ("window".equals(properties.getProperty("templatePath"))) ?
                    properties.getProperty("windowTemplatePath") :
                    properties.getProperty("linuxTemplatePath");
            exceInputStream = new FileInputStream(new File(filePath));
        } catch (IOException e) {
            log.info("加载配置文件读取字节流异常："+e.getMessage());
        }
    }


    public static void main(String[] args) {
        testClassLoad load = new testClassLoad();
        log.info("类创建成功1...");
        testClassLoad load2 = new testClassLoad();
        log.info("类创建成功2...");
    }

}
