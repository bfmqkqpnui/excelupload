package com.lance.utils;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

/**
 * excel工具类
 *
 * @author
 * @create 2018-02-04 8:43
 **/
public class ExcelUtil {
    public static final String OFFICE_EXCEL_2003_POSTFIX = "xls";
    public static final String OFFICE_EXCEL_2010_POSTFIX = "xlsx";
    public static final String OFFICE_EXCEL_97_2003_POSTFIX = "csv";
    public static final String EMPTY = "";
    public static final String POINT = ".";
    public static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

    /**
     * 获得path的后缀名
     *
     * @param path
     * @return
     */
    public static String getPostfix(String path) {
        String postfix = null;
        if (StringUtils.isBlank(path)) {
            postfix = EMPTY;
        }
        if (path.contains(POINT)) {
            postfix = path.substring(path.lastIndexOf(POINT) + 1, path.length());
        }
        return postfix;
    }

    /**
     * 读取excel数据
     *
     * @param hssfCell
     * @return
     */
    public static String getHValue(HSSFCell hssfCell) {
        DecimalFormat df = new DecimalFormat("#");
        if (hssfCell == null) {
            return "";
        }

        switch (hssfCell.getCellType()) {
            case HSSFCell.CELL_TYPE_NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(hssfCell)) {
                    return sdf.format(HSSFDateUtil.getJavaDate(hssfCell.getNumericCellValue()));
                }
                return df.format(hssfCell.getNumericCellValue());
            case HSSFCell.CELL_TYPE_BLANK:
                return "";
            case HSSFCell.CELL_TYPE_STRING:
                return hssfCell.getStringCellValue();
            case HSSFCell.CELL_TYPE_FORMULA:
                return hssfCell.getCellFormula();
            default:
                return hssfCell.getStringCellValue();
        }
    }

    /**
     * 读取excel数据
     *
     * @param xssfCell
     * @return
     */
    public static String getXValue(XSSFCell xssfCell) {
        DecimalFormat df = new DecimalFormat("#");
        if (xssfCell == null) {
            return "";
        }

        switch (xssfCell.getCellType()) {
            case XSSFCell.CELL_TYPE_NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(xssfCell)) {
                    return sdf.format(HSSFDateUtil.getJavaDate(xssfCell.getNumericCellValue()));
                }
                return df.format(xssfCell.getNumericCellValue());
            case XSSFCell.CELL_TYPE_BLANK:
                return "";
            case HSSFCell.CELL_TYPE_STRING:
                return xssfCell.getStringCellValue();
            case HSSFCell.CELL_TYPE_FORMULA:
                return xssfCell.getCellFormula();
            default:
                return xssfCell.getStringCellValue();
        }
    }
}
