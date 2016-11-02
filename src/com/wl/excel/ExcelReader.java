package com.wl.excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

public class ExcelReader {
	
	/**
	 * 获取整个列表
	 * @param path
	 * @return
	 */
	public List<Commodity> readFromExcel(String path) {
		InputStream is = null;
		Workbook workbook = null;
		try {
			is = new FileInputStream(path);
//			try {
//				workbook = new HSSFWorkbook(is);
//				return parseFromHssf((HSSFWorkbook) workbook);
//			} catch (OfficeXmlFileException e) {
				workbook = new XSSFWorkbook(is);
				return parseFromXssf((XSSFWorkbook) workbook);
//			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if(is != null)
					is.close();
				if(workbook != null)
					workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return null;
	}
	
//	private List<Commodity> parseFromHssf(HSSFWorkbook hssfWorkbook) {
//		Commodity commodity;
//		List<Commodity> list = new ArrayList<>();
//		for (int i = 0; i < hssfWorkbook.getNumberOfSheets(); i++) {
//			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(i);
//			if(hssfSheet != null) {
//				for(int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum ++) {
//					HSSFRow hssfRow = hssfSheet.getRow(rowNum);
//					if(hssfRow != null) {
//						commodity = new Commodity();
//						commodity.setDatetime(getValue(hssfRow.getCell(0)));
//						commodity.setType((int)Double.parseDouble(getValue(hssfRow.getCell(1))));
//						commodity.setName(getValue(hssfRow.getCell(2)));
//						commodity.setPrice(Double.parseDouble(getValue(hssfRow.getCell(3))));
//						commodity.setCount((int)Double.parseDouble(getValue(hssfRow.getCell(4))));
//						commodity.setAssistant(getValue(hssfRow.getCell(5)));
//						commodity.setPercentage(Double.parseDouble(getValue(hssfRow.getCell(6))));
//						list.add(commodity);
//					}
//				}
//			}
//		}
//
//		return list;
//	}
//
	private List<Commodity> parseFromXssf(XSSFWorkbook hssfWorkbook) {
		Commodity commodity;
		List<Commodity> list = new ArrayList<>();
		for (int i = 0; i < hssfWorkbook.getNumberOfSheets(); i++) {
			XSSFSheet hssfSheet = hssfWorkbook.getSheetAt(i);
			if(hssfSheet != null) {
                String datetime = null;
				for(int rowNum = 2; rowNum <= hssfSheet.getLastRowNum(); rowNum ++) {
					XSSFRow hssfRow = hssfSheet.getRow(rowNum);
					if(hssfRow != null) {
						commodity = new Commodity();
                        String vdate = getValue(hssfRow.getCell(0));
                        if(vdate == null || "".equals(vdate)) {
                            commodity.setDatetime(datetime);
                        } else {
                            datetime = vdate;
                            commodity.setDatetime(datetime);
                        }


						String vName = getValue(hssfRow.getCell(2));
						if(vName == null || "".equals(vName)) {
							break;
						}
						commodity.setName(vName);
                        commodity.setType((int)Double.parseDouble(getValue(hssfRow.getCell(1))));
						commodity.setPrice(Double.parseDouble(getValue(hssfRow.getCell(4))));
						commodity.setCount((int)Double.parseDouble(getValue(hssfRow.getCell(3))));
						commodity.setAssistant(getValue(hssfRow.getCell(7)));
						String vPer = getValue(hssfRow.getCell(8));
						if(vPer == null || "".equals(vPer))
							vPer = "0";
						commodity.setPercentage(Double.parseDouble(vPer));
						list.add(commodity);
					}
				}
			}
		}
		
		return list;
	}
	
	private String getValue(HSSFCell hssfCell) {
		if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
			// 返回布尔类型的值
			return String.valueOf(hssfCell.getBooleanCellValue());
		} else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
			// 返回数值类型的值
			return String.valueOf(hssfCell.getNumericCellValue());
		} else {
			// 返回字符串类型的值
			return String.valueOf(hssfCell.getStringCellValue());
		}
	}
	
	private String getValue(XSSFCell cell) {
		if (cell.getCellType() == cell.CELL_TYPE_BOOLEAN) {
			// 返回布尔类型的值
			return String.valueOf(cell.getBooleanCellValue());
		} else if (cell.getCellType() == cell.CELL_TYPE_NUMERIC) {
			// 返回数值类型的值
			 short format = cell.getCellStyle().getDataFormat();  
		    SimpleDateFormat sdf = null;  
		    String result = null;
		    if(format == 14 || format == 31 || format == 57 || format == 58 || format == 165 || format == 59){
		        //日期  
		        sdf = new SimpleDateFormat("yyyy-MM-dd");  
		        double value = cell.getNumericCellValue();  
			    Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(value);  
			    result = sdf.format(date);
		    }else if (format == 20 || format == 32) {  
		        //时间  
		        sdf = new SimpleDateFormat("HH:mm");
		        double value = cell.getNumericCellValue();  
			    Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(value);  
			    result = sdf.format(date);
		    }  else {
		    	result = String.valueOf(cell.getNumericCellValue());
		    }
		    
			return result;
		} else {
			// 返回字符串类型的值
			return String.valueOf(cell.getStringCellValue());
		}
	}
	
	/**
	 * 获取每种类型的台数
	 * @param type
	 * @param list
	 * @return
	 */
	public static int getCountByType(int type, List<Commodity> list) {
		if(list == null)
			return 0;
		int count = 0;
		for (int i = 0; i < list.size(); i++) {
			if(list.get(i).getType() == type) {
				count += list.get(i).getCount();
			}
		}
		
		return count;
	}
	
	/**
	 * 按人分类
	 * @param list
	 * @return
	 */
	public static HashMap<String, List<Commodity>> groupByPerson(List<Commodity> list) {
		if(list == null)
			return null;
		HashMap<String, List<Commodity>> map = new HashMap<>();
		Commodity commodity;
		for(int i=0;i< list.size();i++) {
			commodity = list.get(i);
			List<Commodity> comms = map.get(commodity.getAssistant());
			if(comms == null) {
				comms = new ArrayList<>();
				map.put(commodity.getAssistant(), comms);
			}
			comms.add(commodity);
		}
		
		return map;
	}
	
	/**
	 * 获取每种类型的总额
	 * @param type
	 * @param list
	 * @return
	 */
	public static int getSumByType(int type, List<Commodity> list) {
		if(list == null)
			return 0;
		int sum = 0;
		for (int i = 0; i < list.size(); i++) {
			if(list.get(i).getType() == type) {
				sum += list.get(i).getPrice();
			}
		}
		
		return sum;
	}
	
	public static int getSum(List<Commodity> list) {
		if(list == null)
			return 0;
		int sum = 0;
		for (int i = 0; i < list.size(); i++) {
			sum += list.get(i).getPrice();
		}
		
		return sum;
	}
	
	public static int getPercentage(List<Commodity> list) {
		if(list == null)
			return 0;
		int sum = 0;
		for (int i = 0; i < list.size(); i++) {
			sum += list.get(i).getPercentage();
		}
		
		return sum;
	}
	
	/**
	 * 获取组件总数
	 * @param list
	 * @return
	 */
	public static int getComponentCount(List<Commodity> list) {
		if(list == null)
			return 0;
		int count = 0;
		for (int i = 0; i < list.size(); i++) {
			if(list.get(i).getType() == Commodity.TYPE_COMPONENT
					|| list.get(i).getType() == Commodity.TYPE_DATALINE) {
				count += list.get(i).getCount();
			}
		}
		
		return count;
	}
	
	/**
	 * 获取组件总额
	 * @param list
	 * @return
	 */
	public static int getComponentSum(List<Commodity> list) {
		if(list == null)
			return 0;
		int sum = 0;
		for (int i = 0; i < list.size(); i++) {
			if(list.get(i).getType() == Commodity.TYPE_COMPONENT
					|| list.get(i).getType() == Commodity.TYPE_DATALINE) {
				sum += list.get(i).getPrice();
			}
		}
		
		return sum;
	}
	
}
