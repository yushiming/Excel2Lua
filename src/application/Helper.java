package application;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class Helper {

	public static void excel2Lua(String inFile, String outFile, String fileName) { //String file = "ItemComposeList.xlsx";
		// TODO Auto-generated method stub
		
		Workbook wb = getWorkBook(inFile);
		if (null == wb){
			System.out.print("wb == null");
		}
		
		List<Map<String,String>> dataList = readExcel(wb);
		writeConsole(dataList);
		writeLua(dataList, outFile, fileName);
	}
	
	
    public static Workbook getWorkBook(String filePath){
 
    	if(null == filePath){
    		return null;
    	}
    	File finalXlsxFile = new File(filePath);
    	String extString = filePath.substring(filePath.lastIndexOf("."));
    	FileInputStream is = null;
    	try {
    		is = new FileInputStream(finalXlsxFile);
    		if (".xls".equals(extString)){
    			return new HSSFWorkbook(is);
    		}else if (".xlsx".equals(extString)){
    			return new XSSFWorkbook(is);
    		}else{
    			return null;
    		}	
    	}catch(FileNotFoundException e){
    		e.printStackTrace();
    	}catch(IOException e){
    		e.printStackTrace();
    	}
    	return null;
    }

	//读取excel
	public static List<Map<String,String>> readExcel(Workbook wb){
		if(wb == null){
			return new ArrayList<Map<String,String>>();	
		}
		
		//用来存放表中数据
		List<Map<String,String>> datalist = new ArrayList<Map<String,String>>();
		
		Sheet sheet = wb.getSheetAt(0);//获取第一个sheet
		if(sheet == null){
			return new ArrayList<Map<String,String>>();	
		}
		int rowNum = sheet.getPhysicalNumberOfRows();//获取最大行数
		
		//第一行为中文名字
		//读取第二行列名 英文名字
		Row rowFirst = sheet.getRow(1);
		int colNum = rowFirst.getPhysicalNumberOfCells();//获取最大列数
		
		ArrayList<String> colNamelist = new ArrayList<String>();
		for(int k = 0; k < colNum; k++){
			Cell cell = rowFirst.getCell(k);
			if (cell != null) 
				colNamelist.add(cell.toString());
			else
				colNamelist.add("nil");
		}
		
		//读取剩下的行
		for(int i = 2; i < rowNum; i++){
			Row row = sheet.getRow(i); //取行
			if(null == row)
				continue;
			
			Map<String,String> rowMap = new LinkedHashMap<String,String>();
			
			for(int j = 0; j < colNum; j++){
				Cell cell = row.getCell(j);
				//if(cell == null)
				//	continue;
					
				String cellData = getValue(cell).toString();
				rowMap.put(colNamelist.get(j), cellData);
			}
			datalist.add(rowMap);
		}
		
		return datalist;
	}
	
	//写到Lua文件
	public static void writeLua(List<Map<String,String>> dataList, String outfile, String fileName){
		StringBuilder content = new StringBuilder("");  
		content.append(fileName + " = {\n");
		
		for(Map<String, String> data : dataList){
			content.append("    {\n");
			for(Map.Entry<String, String> entry : data.entrySet()){
				String mapKey = entry.getKey();
				String mapValue = entry.getValue();
				content.append("        " + "[\"" + mapKey + "\"]" + " = " + mapValue + "," + "\n");
			}
			content.append("    },\n");
		}
		content.append("}");
		try {
		   File file = new File(outfile);

		   // if file doesnt exists, then create it
		   if (!file.exists()) {
		    file.createNewFile();
		   }

		   //FileWriter fw = new FileWriter(file.getAbsoluteFile());
		   
		   PrintWriter out = new PrintWriter(new BufferedWriter(new OutputStreamWriter(new FileOutputStream(file.getAbsoluteFile()),"utf-8")));
		   //BufferedWriter bw = new BufferedWriter(fw);
		   out.write(content.toString());
		   out.close();

		   System.out.println("Done");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	//写到控制台
	public static void writeConsole(List<Map<String,String>> dataList){
		for(Map<String, String> data : dataList){
			for(Map.Entry<String, String> entry : data.entrySet()){
				String mapKey = entry.getKey();
				String mapValue = entry.getValue();
				System.out.print(mapKey + " : " + mapValue);
				System.out.print("\n");
			}
			System.out.print("\n\n");
		}
	}
	
	private static Object getValue(Cell cell) {
    	Object obj = null;
    	
    	if (cell == null){
    		obj = "nil";
    		return obj;
    	}
    	
    	switch (cell.getCellType()) {
	        case Cell.CELL_TYPE_BOOLEAN:
	            obj = cell.getBooleanCellValue(); 
	            break;
	        case Cell.CELL_TYPE_ERROR:
	            obj = cell.getErrorCellValue(); 
	            break;
	        case Cell.CELL_TYPE_NUMERIC:
	            obj = (int)(cell.getNumericCellValue());
	            break;
	        case Cell.CELL_TYPE_STRING:
	            obj = "\"" + cell.getStringCellValue() + "\""; 
	            break;
	        case Cell.CELL_TYPE_BLANK:
	            obj = "nil";// "\"" + cell.getStringCellValue() + "\""; 
	            break;
	        default:
	            break;
    	}
    	
    	if (obj == null){
    		obj = "nil";
    	}
    	
    	return obj;
	}
	
	/*
	  public static final int CELL_TYPE_NUMERIC = 0;
	  
	  // Field descriptor #4 I
	  public static final int CELL_TYPE_STRING = 1;
	  
	  // Field descriptor #4 I
	  public static final int CELL_TYPE_FORMULA = 2;
	  
	  // Field descriptor #4 I
	  public static final int CELL_TYPE_BLANK = 3;
	  
	  // Field descriptor #4 I
	  public static final int CELL_TYPE_BOOLEAN = 4;
	  
	  // Field descriptor #4 I
	  public static final int CELL_TYPE_ERROR = 5;
	*/

}
