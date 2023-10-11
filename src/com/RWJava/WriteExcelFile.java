package com.RWJava;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelFile {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		//Blank Workbook
		XSSFWorkbook workbook=new XSSFWorkbook();
		
		//Create Excel Sheet
		XSSFSheet sampleSheet = workbook.createSheet("SampleSheet");
		
		//creating the data
		Map<String,Object[]> dataSet=new TreeMap<String, Object[]>();
		
		dataSet.put("1", new Object[] {"ID","NAME","Company"});
		dataSet.put("2", new Object[] {"1","James","Pertline Inc"});
		dataSet.put("3", new Object[] {"2","Maria","Sumologic Inc"});
		dataSet.put("4", new Object[] {"3","Peter","Siemens corp."});
		dataSet.put("5", new Object[] {"4","Julia","Google Inc"});
		dataSet.put("6", new Object[] {"5","Ajay","Facebook Inc"});
		
		//Iterate Over the Data
		Set<String> set=dataSet.keySet();
		int rownum=0;
		
		for(String key:set) {
			//create row
			Row row=sampleSheet.createRow(rownum++);
			
			Object[] data=dataSet.get(key);
			
			int cellnum=0;
			for(Object value: data) {
				//create Column
				Cell cell =row.createCell(cellnum++);
				
				if(value instanceof String ) {
					cell.setCellValue((String) value);
				}else if(value instanceof Integer) {
					cell.setCellValue((Integer)value);
				}
			}
		}
		//Write Down File on HardDisk
		try {
			FileOutputStream writeFile=new FileOutputStream("sampleSheet.xlsx");
			
			workbook.write(writeFile);
			writeFile.close();
			System.out.println("Sample Excel file is being create succesfully");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
