package com.mycompany.app;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;


public class ExcelReader {

	 public static String sExcelPath = System.getProperty("user.dir")+ "/UserData.xls";

	 public static void main(String[] argv) throws Exception {
		 
//		 //	 read excel worksheet and print on console	   
//		   readExcelAndPrintOnCOnsole();
//		   
//		 updateRecordInExcelSheet();
//		   	  
//
//	 readDataFromMysqlDatabaseAndPrintToConsole();
//	
	 readDataFromMysqlDatabaseAndWriteToExcel();
	 
	 }
	 

	
	 private static void readDataFromMysqlDatabaseAndWriteToExcel() throws Exception{
		 GenericFunctions genericFunctions =  new GenericFunctions();
		 	Connection con ;
		    ResultSet rs;
		 	
		    Connection excelConnection ;
		    ResultSet excelrs;
		    String excelQuery = "";
		    
		   
		
		    String sQuery = "Select * from app_user";
		   con = genericFunctions.GetDBConnection();
		   Statement stmt = con.createStatement();
		   
		   rs = stmt.executeQuery(sQuery);	   

		   
		   
		   excelConnection = genericFunctions.GetExcelConnection(sExcelPath, "false");
		   Statement excelStmt = excelConnection.createStatement();
		   
		 
		      
		 while (rs.next()) {
		        String uname = rs.getString(2);
		        String pwd = rs.getString(3);
		        int id = rs.getInt(1);
		        
		        excelQuery = "insert into [Users$](id,UserName,password) values ('"+id+ "','"+uname+"','"+pwd+"')";
		        String insertQuery = "";
		        excelStmt.executeUpdate(excelQuery);
				 excelQuery = "select count(id) as cnt from [Users$] ";

				 
//				 excelrs =   excelStmt.executeQuery(excelQuery);
//				 if (excelrs.getInt("cnt")>0) {
//					 insertQuery = "insert into [insertEntries$](inserted_row_id) values ('"+ excelrs.getInt("cnt")+1 + "')";
//					 excelStmt.executeUpdate(insertQuery);
//					 System.out.println("###### "+excelrs.getString(1));
//				 }

		        
		      }

		 excelStmt.close();
		 excelConnection.close();
		   
		 	  rs.close();
		 	  stmt.close();
		 	  con.close();
		 	  
		 	  System.out.println("--------------------------------------------------------------------------");

		 
	}



	private static void readDataFromMysqlDatabaseAndPrintToConsole() throws Exception {
		 GenericFunctions genericFunctions =  new GenericFunctions();
		 	Connection con ;
		    ResultSet rs;
		 	String sQuery = "Select * from app_user";
		   con = genericFunctions.GetDBConnection();
		   
		   
		   Statement stmt = con.createStatement();
		   
		   rs = stmt.executeQuery(sQuery);
		 
		   System.out.println("ID \t\t   UserName \t \t password ");
		      
		 while (rs.next()) {
		        String uname = rs.getString(2);
		        String pwd = rs.getString(3);
		        int id = rs.getInt(1);
		        
		                
		      }
		 	  rs.close();
		 	  stmt.close();
		 	  con.close();
		 	  
		 	  System.out.println("--------------------------------------------------------------------------");

		 
	}



	private static void updateRecordInExcelSheet() throws Exception{
		 	GenericFunctions genericFunctions =  new GenericFunctions();
		 	Connection con ;
		    ResultSet rs;
		    String sQuery = "update [Users$] set UserName='Updated Name' where id=1";
		    
		   con = genericFunctions.GetExcelConnection(sExcelPath, "false");
		   
		   Statement stmt = con.createStatement();
		
		   stmt.executeUpdate(sQuery);
		   
		   System.out.println(stmt.executeUpdate(sQuery));
		   stmt.close();
		   con.close();
		   
	}



	public static void readExcelAndPrintOnCOnsole() throws Exception
	 {
		 GenericFunctions genericFunctions =  new GenericFunctions();
		 	Connection con ;
		    ResultSet rs;
		 	String sQuery = "Select * from [Users$]";
		   con = genericFunctions.GetExcelConnection(sExcelPath, "false");
		   
		   
		   Statement stmt = con.createStatement();
		   
		   rs = stmt.executeQuery(sQuery);
		 
		   System.out.println("ID \t\t   UserName \t \t password ");
		      
		 while (rs.next()) {
		        String uname = rs.getString(2);
		        String pwd = rs.getString(3);
		        int id = rs.getInt(1);
				 System.out.println(">>>>"+ rs.getRow());		

		        
		        System.out.println(id + " \t\t" + uname+ "  \t\t" + pwd);
		      }

		 	  rs.close();
		 	  stmt.close();
		 	  con.close();
		 	  
		 	  System.out.println("--------------------------------------------------------------------------");
	 }
	 
}
