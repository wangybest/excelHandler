package com.zjut.excel.controller;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import com.zjut.excel.vo.MaterialModel;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @author wangyong
 * @date 2021/5/8  15:27
 */
@RestController
@RequestMapping("/rest/excel")
public class ExcelHandler {

    private static Integer rowNumber = 3;

    @PostMapping("/upload")
    public void handleExcel(MultipartFile file,String supplier,String projectName, HttpServletResponse httpServletResponse) throws Exception {
        AtomicInteger i = new AtomicInteger(1);
        List<MaterialModel> materialModels = ExcelImportUtil.importExcel(file.getInputStream(), MaterialModel.class, new ImportParams());
        materialModels.forEach(item->{
            item.setSupplier(supplier);
            item.setProjectName(projectName);
            item.setSerialNumber(String.valueOf(i.getAndIncrement()));
        });

        int rowNo = 0;
        int columnNo = 0;
        int material = rowNumber;
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        XSSFSheet xssfSheet = xssfWorkbook.createSheet();
        CellStyle style = xssfWorkbook.createCellStyle();
        //下边框
        style.setBorderBottom(BorderStyle.THIN);
        //左边框
        style.setBorderLeft(BorderStyle.THIN);
        //上边框
        style.setBorderTop(BorderStyle.THIN);
        //右边框
        style.setBorderRight(BorderStyle.THIN);
        int width = rowNumber * 3 - 1;
        int unitHeight = materialModels.size() / rowNumber + (materialModels.size() % rowNumber > 0 ? 1 : 0);
        int height = unitHeight * (MaterialModel.class.getDeclaredFields().length + 1) - 1;
        for (int j = 0; j < height; j++) {
            XSSFRow xssfRow = xssfSheet.createRow(j);
            for (int k = 0; k < width; k++) {
                xssfRow.createCell(k);
            }
        }
        for (MaterialModel materialModel : materialModels) {
            Field[] fields = materialModel.getClass().getDeclaredFields();
            Excel[] excels = new Excel[fields.length];
            for (int i1 = 0; i1 < fields.length; i1++) {
                Excel excel = fields[i1].getDeclaredAnnotation(Excel.class);
                excels[i1] = excel;
            }
            for (int x = 0; x < materialModel.getClass().getDeclaredFields().length; x++) {
                xssfSheet.getRow(rowNo).getCell(columnNo).setCellValue(excels[x].name());
                fields[x].setAccessible(true);
                xssfSheet.getRow(rowNo++).getCell(columnNo + 1).setCellValue((String) fields[x].get(materialModel));
            }
            rowNo++;
            columnNo = columnNo + 3;
            material = material == 3 ? 1 : material + 1;
            if (material != rowNumber) {
                rowNo = rowNo - materialModel.getClass().getDeclaredFields().length - 1;
            } else {
                columnNo = columnNo - rowNumber * 3;
            }
        }
        for (int j = 0; j < height; j++) {
            XSSFRow xssfRow = xssfSheet.getRow(j);
            xssfSheet.setColumnWidth(j,5000);
            for (int k = 0; k < width; k++) {
                XSSFCell cell = xssfRow.getCell(k);
                if (StringUtils.isNotBlank(cell.getStringCellValue())) {
                    cell.setCellStyle(style);
                }
            }
        }
        httpServletResponse.setContentType("application/vnd.ms-excel");
        httpServletResponse.setCharacterEncoding("utf-8");
        httpServletResponse.setHeader("Content-disposition", "attachment;filename*=utf-8''" + "workbook.xlsx");
        xssfWorkbook.write(httpServletResponse.getOutputStream());
    }
}
