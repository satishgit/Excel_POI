package com.mycompany.app;
import java.sql.Connection;
import java.sql.DriverManager;


public class GenericFunctions {

	
	String sDataBaseType = "MySQL";
	String sDBHostName = "127.0.0.1";
	String sDBPortNumber ="3306";
	String sDBSID = "test";
	String sDBUserName = "root";
	String sDBPassword = "";
	Connection oConnection;
	
	 public static String sExcelPath = System.getProperty("user.dir")+ "/UserData.xls";

	
	/*************************************************
	 * *************************************************/
	/**
	 * Description   : Returns a connection to the application database.
	 * @param sDBType : Type of database to connect to (optional)
	 * @param sPath : DB file path (optional)
	 * @return oConnection: Connection object in case of successful connection, null otherwise.
	 * @author Dattatrya More
	 */
	public Connection GetDBConnection() throws Exception 
	{
		String sConnStr = null;
		//Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
		if (sDataBaseType.equalsIgnoreCase("DB2")) 
		{
			//sConnStr = "jdbc:odbc:Driver={IBM DB2 ODBC DRIVER};Database=" + sSID + ";Hostname=" + sHostName + ";Port="+ sPortNumber +";Protocol=TCPIP;Uid=" + sDBUserName + ";Pwd=" + sDBPassword + ";";		
			Class.forName("com.ibm.db2.jcc.DB2Driver");
			sConnStr = "jdbc:db2://" + sDBHostName + ":" + sDBPortNumber + "/" + sDBSID + "";
		}
		else if (sDataBaseType.equalsIgnoreCase("Oracle")) 
		{
			Class.forName("oracle.jdbc.driver.OracleDriver");
			sConnStr = "jdbc:oracle:thin:@" + sDBHostName + ":" + sDBPortNumber + ":" + sDBSID;   
		}
		else if (sDataBaseType.equalsIgnoreCase("MySQL")) 
		{
			Class.forName("com.mysql.jdbc.Driver");
			sConnStr = "jdbc:mysql://"+ sDBHostName +":"+ sDBPortNumber + "/" + sDBSID + "";
		}
		else
		{
			return null;
		}
		if(null == oConnection)
		{
			oConnection = DriverManager.getConnection(sConnStr,sDBUserName,sDBPassword);
		}
		return oConnection;
	} // End GetDBConnection

	/**************************************************************************************************/
	/**
	 * Description   : Returns a connection to the Excel database.
	 * @param sPath : Excel file path 
	 * @param sReadOnly: false for write access.
	 * @return oConnection: Connection object in case of successful connection, null otherwise.
     */
	public Connection GetExcelConnection(String sPath, String sReadOnly) throws Exception 
	{
		Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
		String sConnStr =
			"jdbc:odbc:Driver={Microsoft Excel Driver (*.xls)};DBQ=" + sPath + ";"
			+ "DriverID=22;READONLY=" + sReadOnly;
		Connection oConnection = DriverManager.getConnection(sConnStr,"","");
		return oConnection;
	} // End GetExcelConnection
	
}
