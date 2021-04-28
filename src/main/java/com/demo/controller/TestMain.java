package com.demo.controller;

import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 直接生成到本地文件
 */
public class TestMain {
    public static void main(String[] args){
        try {
            exportExcwl();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public static void exportExcwl()throws Exception{
        List<Object> list = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            Map<String,Object> data = new HashMap<>();
            data.put("a1", (int)(Math.random()*100) );
            data.put("a2", (int)(Math.random()*100) );
            data.put("a3", (int)(Math.random()*100) );
            data.put("a4", (int)(Math.random()*100) );
            data.put("a5", (int)(Math.random()*100) );
            list.add(data);
        }
        // 表格使用的数据
        Map map = new HashMap();
        map.put("data",list);
        map.put("title","java基于模板导出excel表格");
        map.put("val","演示合并单元格的数据显示");
        // 获取模板文件
        //InputStream is = this.getClass().getClassLoader().getResourceAsStream("x1.xls");
        InputStream is = new BufferedInputStream(new FileInputStream(new File("D:\\data\\template.xlsx")));
        // 实例化 XLSTransformer 对象
        XLSTransformer xlsTransformer = new XLSTransformer();
        // 获取 Workbook ，传入 模板 和 数据
        Workbook workbook =  xlsTransformer.transformXLS(is,map);
        // 写出文件
        OutputStream os = new BufferedOutputStream(new FileOutputStream("D:\\data\\temp.xlsx"));
        // 输出
        workbook.write(os);
        // 关闭和刷新管道，不然可能会出现表格数据不齐，打不开之类的问题
        is.close();
        os.flush();
        os.close();
    }
}
