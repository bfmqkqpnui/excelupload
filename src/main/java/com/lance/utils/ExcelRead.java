package com.lance.utils;

import com.opencsv.CSVReader;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;
import sun.plugin2.ipc.unix.DomainSocketNamedPipe;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * 读取excel
 *
 * @author
 * @create 2018-02-04 8:41
 **/
public class ExcelRead {
    //sheet中总行数
    public int totalRows;
    //每一行总单元格数
    public static int totalCells;

    /**
     * 读取excel的数据
     *
     * @param file
     * @return
     * @throws IOException
     */
    public List<List<String>> readExcel(MultipartFile file) throws IOException {
        if (file == null || ExcelUtil.EMPTY.equals(file.getOriginalFilename().trim())) {
            return null;
        } else {
            String postfix = ExcelUtil.getPostfix(file.getOriginalFilename());
            if (!ExcelUtil.EMPTY.equals(postfix)) {
                if (ExcelUtil.OFFICE_EXCEL_2003_POSTFIX.equalsIgnoreCase(postfix)) {
                    return readXls(file);
                } else if (ExcelUtil.OFFICE_EXCEL_2010_POSTFIX.equalsIgnoreCase(postfix)) {
                    return readXlsx(file);
                } else if (ExcelUtil.OFFICE_EXCEL_97_2003_POSTFIX.equalsIgnoreCase(postfix)) {
                    return readCsv(file);
                } else {
                    return null;
                }
            }
        }
        return null;
    }

    /**
     * 读取excel-2010  .xlsx
     *
     * @param file
     * @return
     */
    public List<List<String>> readXlsx(MultipartFile file) {
        List<List<String>> list = new ArrayList<List<String>>();
        // IO流读取文件
        InputStream input = null;
        XSSFWorkbook wb = null;
        ArrayList<String> rowList = null;
        try {
            input = file.getInputStream();
            // 创建文档
            wb = new XSSFWorkbook(input);
            //读取sheet(页)
            for (int numSheet = 0; numSheet < wb.getNumberOfSheets(); numSheet++) {
                XSSFSheet xssfSheet = wb.getSheetAt(numSheet);
                if (xssfSheet == null) {
                    continue;
                }
                totalRows = xssfSheet.getLastRowNum();
                //读取Row,从第二行开始
                for (int rowNum = 1; rowNum <= totalRows; rowNum++) {
                    XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                    if (xssfRow != null) {
                        rowList = new ArrayList<String>();
                        totalCells = xssfRow.getLastCellNum();
                        //读取列，从第一列开始
                        for (int c = 0; c <= totalCells + 1; c++) {
                            XSSFCell cell = xssfRow.getCell(c);
                            if (cell == null) {
                                rowList.add(ExcelUtil.EMPTY);
                                continue;
                            }
                            rowList.add(ExcelUtil.getXValue(cell).trim());
                        }
                        list.add(rowList);
                    }
                }
            }
            return list;
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                input.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return null;

    }

    /**
     * 读取excel-2003_2007  .xls
     *
     * @param file
     * @return
     */
    public List<List<String>> readXls(MultipartFile file) {
        List<List<String>> list = new ArrayList<List<String>>();
        // IO流读取文件
        InputStream input = null;
        HSSFWorkbook wb = null;
        List<String> rowList = null;
        try {
            input = file.getInputStream();
            // 创建文档
            wb = new HSSFWorkbook(input);
            //读取sheet(页)
            for (int numSheet = 0; numSheet < wb.getNumberOfSheets(); numSheet++) {
                HSSFSheet hssfSheet = wb.getSheetAt(numSheet);
                if (hssfSheet == null) {
                    continue;
                }
                totalRows = hssfSheet.getLastRowNum();
                //读取Row,从第二行开始
                for (int rowNum = 1; rowNum <= totalRows; rowNum++) {
                    HSSFRow hssfRow = hssfSheet.getRow(rowNum);
                    if (hssfRow != null) {
                        rowList = new ArrayList<String>();
                        totalCells = hssfRow.getLastCellNum();
                        //读取列，从第一列开始
                        for (short c = 0; c <= totalCells + 1; c++) {
                            HSSFCell cell = hssfRow.getCell(c);
                            if (cell == null) {
                                rowList.add(ExcelUtil.EMPTY);
                                continue;
                            }
                            rowList.add(ExcelUtil.getHValue(cell).trim());
                        }
                        list.add(rowList);
                    }
                }
            }
            return list;
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                input.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return null;
    }

    public static List<List<String>> readCsv(MultipartFile file) {
        List<List<String>> list = new ArrayList<List<String>>();
        List<String[]> strArrayList = new ArrayList<String[]>();
        // IO流读取文件
        InputStream input = null;
        CSVReader csvReader = null;
        try {
            input = file.getInputStream();
            InputStreamReader isr = new InputStreamReader(input);
            BufferedReader br = new BufferedReader(isr);
            csvReader = new CSVReader(br);

            String[] rowList;
            int rowNum = 1;
            while ((rowList = csvReader.readNext()) != null) {
                if(rowNum > 1){
                    strArrayList.add(rowList);
                }
                rowNum ++;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        if(CommonUtil.isExist(strArrayList)){
            for(String[] str:strArrayList){
                if(null != str && str.length > 0){
                    List<String> row = Arrays.asList(str);
                    if(CommonUtil.isExist(row)){
                        list.add(row);
                    }
                }
            }
        }

        return list;
    }
}
