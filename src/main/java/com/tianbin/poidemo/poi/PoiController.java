package com.tianbin.poidemo.poi;

import org.apache.poi.hssf.usermodel.*;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import sun.awt.image.IntegerInterleavedRaster;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Auther: T
 * @Date: 2019/10/23 17:10
 * @Description:
 */
@RestController
@RequestMapping("poi")
public class PoiController {

    @RequestMapping("download")
    public void downloadTest(HttpServletRequest request, HttpServletResponse response) throws IOException {
        List<Map<String,String>> rstList=new ArrayList<>();
        for(int i=0;i<10;i++){
            Map<String,String> map=new HashMap<>();
            map.put("name","name_"+i);
            map.put("age",i+"");
            rstList.add(map);
        }
        //新建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //新建表表
        HSSFSheet sheet = workbook.createSheet("导出表");
        //表名
        String fileName = "students" + ".xls";
        //第一行为表头
        String [] headers = {"学号","姓名","身份类型","登录密码"};
        HSSFRow row = sheet.createRow(0);
        for (int i = 0;i<headers.length;i++){
            HSSFCell cell = row.createCell(i);
            HSSFRichTextString text = new HSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        //后面行为数据
        int rowNum = 1;
        for (Map map : rstList){
            HSSFRow row1 = sheet.createRow(rowNum);
            row1.createCell( 0).setCellValue(new HSSFRichTextString((String)map.get("name")));
            row1.createCell(1).setCellValue(new HSSFRichTextString((String) map.get("age")));
            rowNum++;

        }
        response.setContentType("application/octet-stream");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName);
        response.flushBuffer();
        workbook.write(response.getOutputStream());
    }
}
