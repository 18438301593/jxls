package com.demo.controller;

import com.demo.entity.Reports;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 下载导出excel（案例2）
 */
@RestController
public class ExportExcelController3 {

    @GetMapping("exportExcel2")
    public void exportExcel(HttpServletResponse response) throws Exception{
        //下载时的文件名
        response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode("报表3.xlsx" ,"UTF-8"));
        //设置文件类型
        response.setContentType("application/vnd.ms-excel");
        //获取模板
        ClassPathResource classPathResource = new ClassPathResource("template/template3.xlsx");
        InputStream inputStream = classPathResource.getInputStream();
        // 实例化 XLSTransformer 对象
        XLSTransformer xlsTransformer = new XLSTransformer();
        // 获取 Workbook ，传入 模板 和 数据
        Workbook workbook =  xlsTransformer.transformXLS(inputStream,getData());
        // 输出
        workbook.write(response.getOutputStream());
        // 关闭和刷新管道，不然可能会出现表格数据不齐，打不开之类的问题
        inputStream.close();
    }

    //获取数据
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
