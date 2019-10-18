package com.richfit.sod.appserver.utils;

import com.richfit.sod.appserver.common.MessageConfig;
import com.richfit.sod.appserver.entity.SysUser;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;


/**
 * @author ：yangsenhe
 * @date ：Created in 2019/5/30 10:38
 * @description：主要poi导出模板数据（带有下拉框,合并单元格的导出）
 * @modified By：
 * @version: v1.0$
 */
public class ExportExcelTemplateUtil {
    /**
     * @param title               标题,以及sheetName
     * @param headers             表头
     * @param values              表中元素
     * @param categoryToStrings   下拉链表集合
     * @param naturalColumnIndexs 下拉列表列集合
     * @return
     */
    public static HSSFWorkbook getHSSFWorkbook(String title, String headers[], String[][] values, List<String[]> categoryToStrings, int[] naturalColumnIndexs) {
        //创建一个HSSFWorkbook，对应一个Excel文件
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();

        //在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet hssfSheet = hssfWorkbook.createSheet(title);

        //创建标题合并行
        hssfSheet.addMergedRegion(new CellRangeAddress(0, (short) 0, 0, (short) headers.length - 1));

        //设置标题样式
        HSSFCellStyle style = hssfWorkbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);   //设置居中样式
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置标题字体
        Font titleFont = hssfWorkbook.createFont();
        titleFont.setFontHeightInPoints((short) 12);
        style.setFont(titleFont);

        //设置必填项带*
        //设置标题样式
        HSSFCellStyle style2 = hssfWorkbook.createCellStyle();
        Font titleFont2 = hssfWorkbook.createFont();
        titleFont2.setFontHeightInPoints((short) 12);
        titleFont2.setColor(HSSFColor.RED.index);
        style2.setFont(titleFont2);

        //设置值表头样式 设置表头居中
        HSSFCellStyle hssfCellStyle = hssfWorkbook.createCellStyle();
        hssfCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); //水平居中
        hssfCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);   //设置垂直居中
        hssfCellStyle.setBorderBottom(BorderStyle.THIN);
        hssfCellStyle.setBorderLeft(BorderStyle.THIN);
        hssfCellStyle.setBorderRight(BorderStyle.THIN);
        hssfCellStyle.setBorderTop(BorderStyle.THIN);
        //设置标题字体
        Font hssfFont = hssfWorkbook.createFont();
        hssfFont.setFontHeightInPoints((short) 11);
        hssfCellStyle.setFont(hssfFont);
        // 设置背景色
        hssfCellStyle.setFillForegroundColor(HSSFColor.ROSE.index);
        hssfCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        //设置表内容样式
        //创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style1 = hssfWorkbook.createCellStyle();
        style1.setBorderBottom(BorderStyle.THIN);
        style1.setBorderLeft(BorderStyle.THIN);
        style1.setBorderRight(BorderStyle.THIN);
        style1.setBorderTop(BorderStyle.THIN);

        //产生标题行
        HSSFRow hssfRow = hssfSheet.createRow(0);
        hssfRow.setHeight((short) 500);
        HSSFCell cell = hssfRow.createCell(0);
        cell.setCellValue(title);
        cell.setCellStyle(style);


        //产生表头
        HSSFRow row1 = hssfSheet.createRow(1);
        row1.setHeight((short) 500);
        for (int i = 0; i < headers.length; i++) {
            HSSFRichTextString richString = new HSSFRichTextString(headers[i]);
            richString.applyFont(0, 1, titleFont2);
            //设置列宽
            hssfSheet.setColumnWidth(i, 7000);
            HSSFCell hssfCell = row1.createCell(i);
            hssfCell.setCellValue(richString);
            hssfCell.setCellStyle(hssfCellStyle);
        }

        //创建内容
        if (Objects.nonNull(values) && values != null) {
            for (int i = 0; i < values.length; i++) {
                row1 = hssfSheet.createRow(i + 2);
                for (int j = 0; j < values[i].length; j++) {
                    //将内容按顺序赋给对应列对象
                    HSSFCell hssfCell = row1.createCell(j);
                    hssfCell.setCellValue(values[i][j]);
                    hssfCell.setCellStyle(style1);
                }
            }
        }

        //创建下拉框
        Integer endRow = 999;
        String allCell[] = {"A","B","C","D","E","F","G","F","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
        if (categoryToStrings != null && categoryToStrings.size() > 0) {
            //储存所有行
            List<Row> allRow = new ArrayList();
            //储存最长的列的值
            Integer maxRow = 0;
            //单个下拉框
            if (categoryToStrings.size()==1){
                maxRow = categoryToStrings.get(0).length;
            }else{
                //（多个下拉框）获取最长的列的值
                maxRow = categoryToStrings.get(0).length;
                for (int i = 1; i < categoryToStrings.size(); i++) {
                    if (maxRow<categoryToStrings.get(i).length){
                        maxRow = categoryToStrings.get(i).length;
                    }
                }
            }
            String sheetName = "hidden";
            Boolean flag = true;
            // 创建隐藏的sheet
            HSSFSheet hidden = hssfWorkbook.createSheet(sheetName);
            //创建最大行*若干列的表格
            for (int i = 0; i < maxRow; i++) {
                Row row = hidden.createRow(i);
                allRow.add(row);
            }
            for(int i = 0; i < categoryToStrings.size(); i++) {
                if(categoryToStrings.get(i) == null || categoryToStrings.get(i).length < 1) {
                    continue;
                }
                // 循环赋值（为了防止下拉框的行数与隐藏域的行数相对应，将隐藏域加到结束行之后）
                for (int j = 0; j < categoryToStrings.get(i).length; j++) {
                    // A1:A代表隐藏域创建第N列createCell(N)时。以A1列开始A行数据获取下拉数组
                    allRow.get(j).createCell(naturalColumnIndexs[i] - 1).setCellValue(categoryToStrings.get(i)[j]);
                }
                /*Name category1Name = hssfWorkbook.createName();
                if (flag) {
                    category1Name.setNameName(sheetName);
                }*/
                DataValidationHelper helper = hidden.getDataValidationHelper();
                // A1:A代表隐藏域创建第?列createCell(?)时.以A1列开始A行数据获取下拉数组
                //category1Name.setRefersToFormula(sheetName + "!A1:A" + (endRow+categoryToStrings.get(i).length));
                // 加载叫做“hidden”这个sheet的数据
                DVConstraint dvConstraint = DVConstraint.createFormulaListConstraint(sheetName + "!$"+allCell[naturalColumnIndexs[i] - 1]+"$1:$"+allCell[naturalColumnIndexs[i] - 1]+"$" +categoryToStrings.get(i).length);
                // 起始行 终止行 起始列 终止列
                CellRangeAddressList addressList = new CellRangeAddressList(2, endRow, naturalColumnIndexs[i] - 1, naturalColumnIndexs[i] - 1);
                // 绑定下拉框和作用区域
                DataValidation validation = helper.createValidation(dvConstraint, addressList);
                // 这个就是隐藏sheet的地方
                hssfWorkbook.setSheetHidden(1, true); // 1隐藏、0显示
                // 对sheet页生效
                hssfSheet.addValidationData(validation);
                flag = false;
            }
        }
        return hssfWorkbook;
    }

    /**
     * 合并单元格的导出
     * @param title 标题,以及sheetName
     * @param MRHeaders  表头合并展示的字段
     * @param index     表头合并从哪行开始 起始位0
     * @param headers  表头
     * @param values  表中元素
     * @return
     */
    public static HSSFWorkbook getMergedRegionHSSFWorkbook(String title,List<String> headers,String[] MRHeaders,int index, List<Map<String,String>> values){
        //创建一个HSSFWorkbook，对应一个Excel文件
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();

        //在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet hssfSheet = hssfWorkbook.createSheet(title);

        //创建标题合并行
        hssfSheet.addMergedRegion(new CellRangeAddress(0,(short)0,0,(short)index+(headers.size()-index)*2 - 1));

        //设置标题样式
        HSSFCellStyle style = hssfWorkbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);   //设置居中样式
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置标题字体
        Font titleFont = hssfWorkbook.createFont();
        titleFont.setFontHeightInPoints((short) 12);
        style.setFont(titleFont);

        //设置值表头样式 设置表头居中
        HSSFCellStyle hssfCellStyle = hssfWorkbook.createCellStyle();
        hssfCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); //水平居中
        hssfCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);   //设置垂直居中
        hssfCellStyle.setBorderBottom(BorderStyle.THIN);
        hssfCellStyle.setBorderLeft(BorderStyle.THIN);
        hssfCellStyle.setBorderRight(BorderStyle.THIN);
        hssfCellStyle.setBorderTop(BorderStyle.THIN);
        //设置标题字体
        Font hssfFont = hssfWorkbook.createFont();
        hssfFont.setFontHeightInPoints((short) 11);
        hssfCellStyle.setFont(hssfFont);
        // 设置背景色
        hssfCellStyle.setFillForegroundColor(HSSFColor.ROSE.index);
        hssfCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        //设置表内容样式
        //创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style1 = hssfWorkbook.createCellStyle();
        style1.setBorderBottom(BorderStyle.THIN);
        style1.setBorderLeft(BorderStyle.THIN);
        style1.setBorderRight(BorderStyle.THIN);
        style1.setBorderTop(BorderStyle.THIN);

        //产生标题行
        HSSFRow hssfRow = hssfSheet.createRow(0);
        hssfRow.setHeight((short)500);
        HSSFCell cell = hssfRow.createCell(0);
        cell.setCellValue(title);
        cell.setCellStyle(style);

        //产生表头
        HSSFRow row1 = hssfSheet.createRow(1);
        HSSFRow row2 = hssfSheet.createRow(2);
        row1.setHeight((short)500);
        for (int i = 0; i < headers.size(); i++) {
            if (i<index){
                //设置列宽
                hssfSheet.setColumnWidth(i,7000);
                HSSFCell hssfCell = row1.createCell(i);
                HSSFCell hssfCel2 = row2.createCell(i);
                hssfCell.setCellValue(headers.get(i));
                hssfCell.setCellStyle(hssfCellStyle);
                hssfCel2.setCellStyle(hssfCellStyle);
                hssfSheet.addMergedRegion(new CellRangeAddress(1,2,i,i));
            }
        }
        int k = index;
        for (int j = index; j <index+(headers.size()-index)*2; j=j+2) {
            //第一行表头
            hssfSheet.setColumnWidth(j,7000);
            HSSFCell hssfCell = row1.createCell(j);
            HSSFCell hssfCel2 = row1.createCell(j+1);
            hssfCell.setCellValue(headers.get(k));
            hssfCell.setCellStyle(hssfCellStyle);
            hssfCel2.setCellStyle(hssfCellStyle);
            hssfSheet.addMergedRegion(new CellRangeAddress(1,1,j,j+1));
            k++;
            //设置列宽
            hssfSheet.setColumnWidth(j,3500);
            hssfSheet.setColumnWidth(j+1,3500);
            //第二行表头
            HSSFCell hssfCell0 = row2.createCell(j);
            HSSFCell hssfCell1 = row2.createCell(j+1);
            hssfCell0.setCellValue(MRHeaders[0]);
            hssfCell1.setCellValue(MRHeaders[1]);
            hssfCell0.setCellStyle(hssfCellStyle);
            hssfCell1.setCellStyle(hssfCellStyle);

        }
        //创建内容
        if (Objects.nonNull(values)&&values!=null){
            for (int i = 0; i <values.size(); i++){
                row1 = hssfSheet.createRow(i+3);
                Map<String, String> maps = values.get(i);
                int j = 0;
                for (String key: maps.keySet()) {
                    //将内容按顺序赋给对应列对象
                    HSSFCell hssfCell = row1.createCell(j);
                    hssfCell.setCellValue(maps.get(key));
                    hssfCell.setCellStyle(style1);
                    j++;
                }
            }
            //获取操作人
            int cellLast = index+(headers.size()-index)*2 - 1;
            int rowLast = values.size()+3;
            HSSFRow row = hssfSheet.createRow(rowLast);
            hssfSheet.setColumnWidth(rowLast,7000);
            HSSFCell hssfCell = row.createCell(cellLast-1);
            HSSFCell hssfCel2 = row.createCell(cellLast);
            //获取登录人
            SysUser user = AuthorityUtil.getCurrentUser();
            hssfCell.setCellValue(MessageConfig.getMessage("OPERATOR")+":"+user.getUserName());
            hssfCell.setCellStyle(style1);
            hssfCel2.setCellStyle(style1);
            hssfSheet.addMergedRegion(new CellRangeAddress(rowLast,rowLast,cellLast-1,cellLast));
        }
        return hssfWorkbook;
    }


    /**
     * 使用已定义的数据源方式设置一个数据验证
     * @param hidden
     * @param validation
     * @param hssfWorkbook
     * @param formulaString data数据
     * @param naturalRowIndex 终止行
     * @param naturalColumnIndex 下拉的列
     * @return
     */
    public DataValidation getDataValidationByFormula(HSSFSheet hidden, HSSFDataValidation validation,
                                                     HSSFWorkbook hssfWorkbook, String[] formulaString,
                                                     int naturalRowIndex, int naturalColumnIndex) {
        // 加载叫做“hidden”这个sheet的数据
        DVConstraint constraint = DVConstraint.createFormulaListConstraint("hidden");
        // 循环赋值（为了防止下拉框的行数与隐藏域的行数相对应，将隐藏域加到结束行之后）
        for (int i = 0; i < formulaString.length; i++) {
            hidden.createRow(i).createCell(naturalColumnIndex - 1).setCellValue(formulaString[i]);
        }
        // 起始行 终止行 起始列 终止列
        CellRangeAddressList addressList = new CellRangeAddressList(2, formulaString.length + 2, naturalColumnIndex - 1, naturalColumnIndex - 1);
        validation = new HSSFDataValidation(addressList, constraint);
        // 这个就是隐藏sheet的地方咯。
        hssfWorkbook.setSheetHidden(1, true); // 1隐藏、0显示
        return validation;
    }
}
