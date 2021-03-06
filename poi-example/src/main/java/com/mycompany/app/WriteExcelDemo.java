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

import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelDemo 
{
	 public static String sExcelPath = System.getProperty("user.dir")+ "/";

	 public static String excelForManualInput = System.getProperty("user.dir")+ "/manual_input.xlsx";

    
	public static void main(String[] args) throws Exception {
		
		String ExcelFileName = "";
		String workSheetToUpdate = "";
		String seqenceNumber = "";
		String updateValue = "";
		
		if(args.length==4)
		{
		
			ExcelFileName = args[0];
			workSheetToUpdate = args[1];
			seqenceNumber = args[2];
			updateValue = args[3];
			
		}else if(args.length==3)
		{
			ExcelFileName = args[0];
			workSheetToUpdate = args[1];
			updateValue = args[2];
				
		}
		
		System.out.println("ExcelFileName: "+ExcelFileName );
		System.out.println("ExcelWorkSheet: "+workSheetToUpdate );
		System.out.println("seqenceNumber: "+seqenceNumber );
		System.out.println("updateValue: "+updateValue );
		
		WriteExcelDemo writeExcelDemo = new WriteExcelDemo();

		Map<String, Object[]> data = writeExcelDemo.readHardCodedData();
		//		Map<String, Object[]> data = writeExcelDemo.readMySQLData();
		
		writeExcelDemo.writeDataToexcelSheet(ExcelFileName,data,workSheetToUpdate,seqenceNumber,updateValue);
		
		
		
		
		
//		
	}
	
	private  Map<String, Object[]>  readHardCodedData() throws Exception { Map<String, Object[]> data = new TreeMap<String, Object[]>();
    data.put("2", new Object[] {1, "Amit", "Shukla"});
    data.put("3", new Object[] {2, "Lokesh", "Gupta"});
    data.put("4", new Object[] {3, "John", "Adwards"});
    data.put("5", new Object[] {4, "Brian", "Schultz"});
    
	return data;
	}
	
	
	
	private void writeDataToexcelSheet(String excelFileName ,Map<String, Object[]> data,String workSheetToUpdate,String seqenceNumber,String updateValue) throws IOException 
	{

		 //Read the spreadsheet that needs to be updated
        FileInputStream inputExcelFile = new FileInputStream(new File(sExcelPath + excelFileName));
       
        XSSFWorkbook workbook = new XSSFWorkbook(inputExcelFile); 
         
        XSSFSheet employeeDataSheet = workbook.getSheet("Employee Data");

        XSSFSheet updateSheet = workbook.getSheet(workSheetToUpdate);

        System.out.println("sheet.getLastRowNum()>>> "+employeeDataSheet.getLastRowNum());
        
        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = employeeDataSheet.getLastRowNum()+1;
        int updateSheetRowNum = updateSheet.getLastRowNum()+1;
        int updateRowNumber = 0;
        
        boolean flag = true;
        for (String key : keyset)
        {
            
        	Row row = employeeDataSheet.createRow(rownum++);
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
            int employeeDataSheetRowNumber = row.getRowNum() + 1;
            if(flag )
            {
            	updateRowNumber = employeeDataSheetRowNumber;
            	flag =false;
            }
            
            
            
            System.out.println("row number>> "+employeeDataSheetRowNumber +" ID:"+objArr[0]);
        }
        
        try
        {
//        	Row row = updateSheet.createRow(updateSheetRowNum);
//        	for(int i=0;i<2;i++)
//        	{
//                Cell cell = row.createCell(i);
//                if(i==0)
//                	cell.setCellValue(updateRowNumber);
//                else
//                	cell.setCellValue("store value with your logic");
//
//        	}
//        	
        	  for(Row row:updateSheet)
              {
        		  if(row.getRowNum()==0)
        			  continue;
        		  if (seqenceNumber == "" || seqenceNumber == null) {
                		Cell valueCell = row.getCell(1);
 						valueCell.setCellValue(updateValue);
				}
                 else
                 {
                	 Cell cell = row.getCell(0);
 					cell.setCellType(cell.CELL_TYPE_STRING);
 					String seqNumber = cell.getStringCellValue();
 					System.out.println("############ " + seqNumber);
 					if (seqNumber.equals(seqenceNumber)) {
 						Cell valueCell = row.getCell(1);
 						valueCell.setCellValue(updateValue);
 					} 
                 }
                   
               }
        	
        	
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(sExcelPath + excelFileName));
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