package com.api.excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public WriteExcel() {

	}

	//–¥»Îxls
    public void write2003(String fileName,Map<Integer,List<String[]>> map) {
        try {
            HSSFWorkbook wb = new HSSFWorkbook();
            for(int sheetnum=0;sheetnum<map.size();sheetnum++){
                HSSFSheet sheet = wb.createSheet(""+sheetnum);
                List<String[]> list = map.get(sheetnum);
                for(int i=0;i<list.size();i++){
                    HSSFRow row = sheet.createRow(i);
                    String[] str = list.get(i);
                    for(int j=0;j<str.length;j++){
                        HSSFCell cell = row.createCell(j);
                        cell.setCellValue(str[j]);    
                    }
                }
            }
            FileOutputStream outputStream = new FileOutputStream(fileName);
            wb.write(outputStream);
            outputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

	
	//–¥»ÎXlsx
    public void write2007Plus(String fileName,Map<Integer,List<String[]>> map) {
        try {
            XSSFWorkbook wb = new XSSFWorkbook();
            for(int sheetnum=0;sheetnum<map.size();sheetnum++){
                XSSFSheet sheet = wb.createSheet(""+sheetnum);
                List<String[]> list = map.get(sheetnum);
                for(int i=0;i<list.size();i++){
                    XSSFRow row = sheet.createRow(i);
                    String[] str = list.get(i);
                    for(int j=0;j<str.length;j++){
                        XSSFCell cell = row.createCell(j);
                        cell.setCellValue(str[j]);
                    }
                }
            }
            FileOutputStream outputStream = new FileOutputStream(fileName);
            wb.write(outputStream);
            outputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

	public static void main(String[] args) {
		new WriteExcel();
	}

}
