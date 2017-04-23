package com.api.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.api.common.Common;
import com.api.util.CommonUtil;

public class ReadExcel {

	public ReadExcel() {
		//assert file exists
		assertFileExists();
		
	}

	private void assertFileExists() {
		if(CommonUtil.exitsFile(Common.TRANSACTION_FILE_2003)){
			System.out.println("2003");
			read2003(Common.TRANSACTION_FILE_2003);
		}else if(CommonUtil.exitsFile(Common.TRANSACTION_FILE_2007_PLUS)){
			System.out.println("2007 + ");
			WriteExcel wE = new WriteExcel();
			wE.write2007Plus(Common.RESULT_FILE_2007_PLUS, read2007Plus(Common.TRANSACTION_FILE_2007_PLUS));
			
			
		}else{
			System.out.println("不存在 '"+ Common.TRANSACTION_FILE_2003 +"' 或者 '"+Common.TRANSACTION_FILE_2007_PLUS+"' 的Excel文件");
		}
	}
	/**
	 * 
	* <p>Description: </p>
	* @author Qi
	* @date 
	* @param transactionFile2003
	 */
	public Map<Integer, List<String []>> read2003(String transactionFile2003) {
		Map<Integer, List<String[]>> map = new HashMap<Integer, List<String[]>>();
        try {
            InputStream is = new FileInputStream(transactionFile2003);
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
            // 循环工作表Sheet  
            for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
                HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
                if (hssfSheet == null) {
                    continue;
                }
                List<String[]> list = new ArrayList<String[]>();
                //总行数:hssfSheet.getLastRowNum()-hssfSheet.getFirstRowNum()+1);
                // 循环行Row
                for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                    HSSFRow hssfRow = hssfSheet.getRow(rowNum);
                    if (hssfRow == null) {
                        continue;
                    }
                    String[] singleRow = new String[hssfRow.getLastCellNum()];
                    //循环row里面的cell单元
                    for(int column=0;column<hssfRow.getLastCellNum();column++){
                        Cell cell = hssfRow.getCell(column,Row.CREATE_NULL_AS_BLANK);
                        switch(cell.getCellType()){
                            case Cell.CELL_TYPE_BLANK:
                                singleRow[column] = "";
                                break;
                            case Cell.CELL_TYPE_BOOLEAN:
                                singleRow[column] = Boolean.toString(cell.getBooleanCellValue());
                                break;
                            case Cell.CELL_TYPE_ERROR:
                                singleRow[column] = "";
                                break;
                            case Cell.CELL_TYPE_FORMULA:
                                cell.setCellType(Cell.CELL_TYPE_STRING);
                                singleRow[column] = cell.getStringCellValue();
                                if (singleRow[column] != null) {
                                    singleRow[column] = singleRow[column].replaceAll("#N/A", "").trim();
                                }
                                break;
                            case Cell.CELL_TYPE_NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    singleRow[column] = String.valueOf(cell.getDateCellValue());
                                } else {
                                    cell.setCellType(Cell.CELL_TYPE_STRING);
                                    String temp = cell.getStringCellValue();
                                    // 判断是否包含小数点，如果不含小数点，则以字符串读取，如果含小数点，则转换为Double类型的字符串
                                    if (temp.indexOf(".") > -1) {
                                        singleRow[column] = String.valueOf(new Double(temp)).trim();
                                    } else {
                                        singleRow[column] = temp.trim();
                                    }
                                }
                                
                                break;
                            case Cell.CELL_TYPE_STRING:
                                singleRow[column] = cell.getStringCellValue().trim();
                                break;
                            default:
                                singleRow[column] = "";
                                break;
                        }
                    }
                   list.add(singleRow);
                }
                map.put(numSheet, list);
            }
            
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
//        display(map);
        return map;
	}
	/**
	 * 
	* <p>Description: </p>
	* @author Qi
	* @date 
	* @param transactionFile2007Plus
	 */
	private Map<Integer, List<String []>> read2007Plus(String transactionFile2007Plus) {
		Map<Integer, List<String[]>> map = new HashMap<Integer, List<String[]>>();

		try {
			InputStream is = new FileInputStream(transactionFile2007Plus);
			XSSFWorkbook workbook = new XSSFWorkbook(is);
			// 循环工作表Sheet
			for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
				XSSFSheet xssfSheet = workbook.getSheetAt(numSheet);
				if (xssfSheet == null) {
                    continue;
                }
				
				List<String[]> list = new ArrayList<String[]>();
				//循环工作表里的Row
				 for (int row=0;row<=xssfSheet.getLastRowNum();row++){
					 XSSFRow xssfRow = xssfSheet.getRow(row);
					 if (xssfRow == null) {
	                        continue;
	                 }
					 String[] singleRow = new String[xssfRow.getLastCellNum()];
					 //循环row里面的cell单元
					 for(int column=0;column<xssfRow.getLastCellNum();column++){
						 Cell cell = xssfRow.getCell(column,Row.CREATE_NULL_AS_BLANK);
						 switch(cell.getCellType()){
						 	case Cell.CELL_TYPE_BLANK:
						 		singleRow[column] = "";
                                break;
						 	case Cell.CELL_TYPE_BOOLEAN:
                                singleRow[column] = Boolean.toString(cell.getBooleanCellValue());
                                break;
						 	case Cell.CELL_TYPE_ERROR:
                                singleRow[column] = "";
                                break;
						 	case Cell.CELL_TYPE_FORMULA:
                                cell.setCellType(Cell.CELL_TYPE_STRING);
                                singleRow[column] = cell.getStringCellValue();
                                if (singleRow[column] != null) {
                                    singleRow[column] = singleRow[column].replaceAll("#N/A", "").trim();
                                }
                                break;
						 	case Cell.CELL_TYPE_NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                	Date date = cell.getDateCellValue();
                                	DateFormat formater = new SimpleDateFormat("yyyy/MM/dd");
                                	String value = formater.format(date);
                                    singleRow[column] = value;
                                } else {
                                    cell.setCellType(Cell.CELL_TYPE_STRING);
                                    String temp = cell.getStringCellValue();
                                    // 判断是否包含小数点，如果不含小数点，则以字符串读取，如果含小数点，则转换为Double类型的字符串
                                    if (temp.indexOf(".") > -1) {
                                        singleRow[column] = String.valueOf(new Double(temp)).trim();
                                    } else {
                                        singleRow[column] = temp.trim();
                                    }
                                }
                                break;
						 	case Cell.CELL_TYPE_STRING:
                                singleRow[column] = cell.getStringCellValue().trim();
                                break;
                            default:
                                singleRow[column] = "";
                                break;
						 }
					 }
					 list.add(singleRow);
				 }
				 map.put(numSheet, list);
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			System.out.println("无法找到:"+transactionFile2007Plus);
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("无法找到:XSSFWorkbook..");
		}

		display(map);
		return map;
	}
	
	private void display(Map<Integer, List<String[]>> map) {
		Set<Integer> keySet = map.keySet();
		List<String[]> list = null;
		String [] tmp = null;
		for (Integer key : keySet) {///sheet key
			list = map.get(key);
			for (int i = 0; list!=null && i < list.size(); i++) {//row loop
				tmp = list.get(i);
				for (int j = 0; tmp!=null && j<tmp.length; j++) {//cell loop
					String string = tmp[j];
					System.out.println("sheet key:"+key + " _ row loop:"+i +" _ cell loop:"+j+"_"+string);
				}
			}
		}
	}

	public static void main(String[] args) {
		new ReadExcel();
	}

}
