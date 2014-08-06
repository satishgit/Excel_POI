1. Unzip poi_example.rar in favorate location
2. execute command 
	mvn eclipse:eclipse
3. Import project into eclipse workspace
4. set M2_repo path
5. execute following command to run sample program.

mvn exec:java -Dexec.mainClass=com.mycompany.app.WriteExcelDemo


above command will read records from database table and after that write into existing excelsheet