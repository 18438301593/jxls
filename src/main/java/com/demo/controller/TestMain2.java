package com.demo.controller;

import com.demo.entity.Reports;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.*;

/**
 * 直接生成到本地文件（案例2）
 */
public class TestMain2 {
    public static void main(String[] args){
        try {
            exportExcwl();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public static void exportExcwl()throws Exception{
        // 获取模板文件
        //InputStream is = this.getClass().getClassLoader().getResourceAsStream("x1.xls");
        InputStream is = new BufferedInputStream(new FileInputStream(new File("D:\\data\\template2.xlsx")));
        // 实例化 XLSTransformer 对象
        XLSTransformer xlsTransformer = new XLSTransformer();
        // 获取 Workbook ，传入 模板 和 数据
        Workbook workbook =  xlsTransformer.transformXLS(is,getData());
        // 写出文件
        OutputStream os = new BufferedOutputStream(new FileOutputStream("D:\\data\\temp2.xlsx"));
        // 输出
        workbook.write(os);
        // 关闭和刷新管道，不然可能会出现表格数据不齐，打不开之类的问题
        is.close();
        os.flush();
        os.close();
    }

    public static Map<String,Object> getData(){
        Map<String,Object> map = new HashMap<>();
        map.put("title","项目报表");
        map.put("count",123);
        map.put("count1",10);
        List<Reports> list1=new ArrayList<>(5);
        list1.add(new Reports(1,"项目名称1","公司名称1","来源1",123.56));
        list1.add(new Reports(2,"项目名称2","公司名称2","来源2",234.22));
        list1.add(new Reports(3,"项目名称3","公司名称3","来源3",234.4));
        list1.add(new Reports(4,"项目名称4","公司名称4","来源4",998.2));
        list1.add(new Reports(5,"项目名称5","公司名称5","来源5",234D));
        map.put("data",list1);

        map.put("count2",14);
        List<Reports> list2=new ArrayList<>(5);
        list2.add(new Reports(1,"项目名称1","公司名称1","来源1",123.56));
        list2.add(new Reports(2,"项目名称2","公司名称2","来源2",234.22));
        list2.add(new Reports(3,"项目名称3","公司名称3","来源3",234.4));
        list2.add(new Reports(4,"项目名称4","公司名称4","来源4",998.9002));
        list2.add(new Reports(5,"项目名称5","公司名称5","来源5",234D));
        map.put("data2",list2);
        return map;
    }
}
