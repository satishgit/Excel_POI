package com.mycompany.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelDemopld 
{
	 public static String sExcelPath = System.getProperty("user.dir")+ "/poi_demo.xlsx";

	 public static String excelForManualInput = System.getProperty("user.dir")+ "/manual_input.xlsx";

    
	public static void main(String[] args) throws Exception {
		WriteExcelDemopld writeExcelDemo = new WriteExcelDemopld();

		Map<String, Object[]> data = writeExcelDemo.readHardCodedData();
		//		Map<String, Object[]> data = writeExcelDemo.readMySQLData();
		
		writeExcelDemo.writeDataToexcelSheet(data);
//		
	}
	
	private  Map<String, Object[]>  readHardCodedData() throws Exception { Map<String, Object[]> data = new TreeMap<String, Object[]>();
    data.put("2", new Object[] {1, "Amit", "Shukla"});
    data.put("3", new Object[] {2, "Lokesh", "Gupta"});
    data.put("4", new Object[] {3, "John", "Adwards"});
    data.put("5", new Object[] {4, "Brian", "Schultz"});
    
	return data;
	}
	
	
	
	private void writeDataToexcelSheet(Map<String, Object[]> data) throws IOException 
	{

		 //Read the spreadsheet that needs to be updated
        FileInputStream input_document = new FileInputStream(new File(sExcelPath));
       
        XSSFWorkbook workbook = new XSSFWorkbook(input_document); 
         
        XSSFSheet sheet = workbook.getSheet("Employee Data");

        XSSFSheet updateSheet = workbook.getSheet("update");

        System.out.println("sheet.getLastRowNum()>>> "+sheet.getLastRowNum());
        
        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = sheet.getLastRowNum()+1;
        int updateSheetRowNum = updateSheet.getLastRowNum()+1;
        int updateRowNumber = 0;
        
        boolean flag = true;
        for (String key : keyset)
        {
            
        	Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               Cell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
            int rowNumber = row.getRowNum() + 1;
            if(flag )
            {
            	updateRowNumber = rowNumber;
            	flag =false;
            }
            
            
            
            System.out.println("row number>> "+rowNumber +" ID:"+objArr[0]);
        }
        
        try
        {
        	Row row = updateSheet.createRow(updateSheetRowNum);
        	for(int i=0;i<2;i++)
        	{
                Cell cell = row.createCell(i);
                if(i==0)
                	cell.setCellValue(updateRowNumber);
                else
                	cell.setCellValue("store value with your logic");

        	}
        	
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(sExcelPath));
            workbook.write(out);
            out.close();
            System.out.println(sExcelPath + " written successfully on disk.");
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
	}
	
	
    private Map<String, Object[]> readMySQLData() throws Exception {
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        GenericFunctions genericFunctions =  new GenericFunctions();
	 	Connection con ;
	    ResultSet rs;
	 	String sQuery = "Select * from employee";
	 	con = genericFunctions.GetDBConnection();
	 	Statement stmt = con.createStatement();
	   
	   rs = stmt.executeQuery(sQuery);
	   
	   
	   while (rs.next()) {
	        int id = rs.getInt(1);
	        String name = rs.getString(2);
	        String lastName = rs.getString(3);
	        
	        System.out.println("ID: "+id +" Name: "+name +"Last Name: "+lastName);
	       
	        data.put("row_"+id, new Object[]{id,name,lastName}); 
	   }
        return data;
	}
    
}