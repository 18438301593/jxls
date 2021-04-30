package com.demo.controller;

import com.demo.entity.Reports;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 直接生成到本地文件（案例2）
 */
public class TestMain3 {
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
        InputStream is = new BufferedInputStream(new FileInputStream(new File("D:\\data\\template3.xlsx")));
        // 实例化 XLSTransformer 对象
        XLSTransformer xlsTransformer = new XLSTransformer();
        // 获取 Workbook ，传入 模板 和 数据
        Workbook workbook =  xlsTransformer.transformXLS(is,getData());
        // 写出文件
        OutputStream os = new BufferedOutputStream(new FileOutputStream("D:\\data\\temp3.xlsx"));
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

        List<Map<String,Object>> dataList=new ArrayList<>();

        Map<String,Object> map1=new HashMap<>();
        List<Reports> list1=new ArrayList<>(5);
        list1.add(new Reports(1,"项目名称1","公司名称1","来源1",123.56));
        list1.add(new Reports(2,"项目名称2","公司名称1","来源2",234.22));
        list1.add(new Reports(3,"项目名称3","公司名称1","来源3",234.4));
        list1.add(new Reports(4,"项目名称4","公司名称1","来源4",998.2));
        list1.add(new Reports(5,"项目名称5","公司名称1","来源5",234D));
        map1.put("data",list1);
        map1.put("num","一");
        map1.put("name","公司名称1");
        map1.put("count","123");
        dataList.add(map1);

        Map<String,Object> map2=new HashMap<>();
        List<Reports> list2=new ArrayList<>(5);
        list2.add(new Reports(1,"项目名称1","公司名称2","来源1",123.56));
        list2.add(new Reports(2,"项目名称2","公司名称2","来源2",234.22));
        list2.add(new Reports(3,"项目名称3","公司名称2","来源3",234.4));
        list2.add(new Reports(4,"项目名称4","公司名称2","来源4",998.2));
        list2.add(new Reports(5,"项目名称5","公司名称2","来源5",234D));
        map2.put("data",list2);
        map2.put("num","二");
        map2.put("name","公司名称2");
        map2.put("count","1232");
        dataList.add(map2);

        map.put("list",dataList);

        return map;
    }
}
