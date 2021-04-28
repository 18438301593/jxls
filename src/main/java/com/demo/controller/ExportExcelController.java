package com.demo.controller;

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
 * 下载导出excel
 */
@RestController
public class ExportExcelController {

    @GetMapping("exportExcel")
    public void exportExcel(HttpServletResponse response) throws Exception{
        //下载时的文件名
        response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode("报表.xlsx" ,"UTF-8"));
        //设置文件类型
        response.setContentType("application/vnd.ms-excel");
        //获取模板
        ClassPathResource classPathResource = new ClassPathResource("template/template.xlsx");
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
    public Map<String,Object> getData(){
        List list = new ArrayList<>();
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
        return map;
    }
}
