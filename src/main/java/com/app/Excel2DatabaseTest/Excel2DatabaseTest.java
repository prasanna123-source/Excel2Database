package com.app.Excel2DatabaseTest;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Excel2DatabaseTest {

	public static void main(String[] args) throws InvalidFormatException {
		// TODO Auto-generated method stub
		String jdbcURL = "jdbc:mysql://localhost:3306/employee";
        String username = "root";
        String password = "admin";
 
      //  String excelFilePath = "Employee_data.xls";
 
        int batchSize = 50;
 
        Connection connection = null;
 
        try {
            long start = System.currentTimeMillis();
             
            FileInputStream inputStream = new FileInputStream("D:\\prasanna_workspace\\demo\\src\\main\\java\\com\\app\\Excel2DatabaseTest\\Students.xlsx");
 
//           Workbook workbook = new XSSFWorkbook(inputStream);
//            workbook = WorkbookFactory.create(inputStream);    
//            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inputStream));
//            HSSFSheet sheet = workbook.getSheetAt(0);
            
            
            Workbook workbook = new XSSFWorkbook(inputStream); 
            org.apache.poi.ss.usermodel.Sheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = firstSheet.iterator();
 
 
            connection = DriverManager.getConnection(jdbcURL, username, password);
            System.out.println("connection created");
            connection.setAutoCommit(false);
  
//            String sql = "INSERT INTO employeedata (empId, empName, gender,bdate,salary) VALUES (?, ?,?,?,?)";
//            PreparedStatement statement = connection.prepareStatement(sql);    
//             
//            int count = 0;
//             
//            rowIterator.next(); // skip the header row
//             
//            while (rowIterator.hasNext()) {
//                Row nextRow = rowIterator.next();
//                Iterator<Cell> cellIterator = nextRow.cellIterator();
// 
//                while (cellIterator.hasNext()) {
//                    Cell nextCell = cellIterator.next();
// 
//                    int columnIndex = nextCell.getColumnIndex();
// 
//                    switch (columnIndex) {
//                    case 0:
//                    	int empId = (int) nextCell.getNumericCellValue();
//                        statement.setInt(1, empId);
//                        break;
//                    case 1:
//                    	String empName = nextCell.getStringCellValue();
//                        statement.setString(2, empName);
//                        break;
//                    case 2:
//                    	String gender = nextCell.getStringCellValue();
//                        statement.setString(3, gender);    
//                    case 3:
//                    	Date bdate = nextCell.getDateCellValue();
//                        statement.setTimestamp(4, new Timestamp(bdate.getTime()));           
//                        
//                    
//                    case 4:
//                        String salary = nextCell.getStringCellValue();
//                        statement.setString(5, salary);                  
// 
//                }
//                 
//                statement.addBatch();
//                 
//                if (count % batchSize == 0) {
//                    statement.executeBatch();
//                }              
// 
//            }
            String sql = "INSERT INTO students (name, enrolled) VALUES (?, ?)";
            PreparedStatement statement = connection.prepareStatement(sql);		
			
            int count = 0;
            
            rowIterator.next(); // skip the header row
            
			while (rowIterator.hasNext()) {
				Row nextRow = rowIterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();

				while (cellIterator.hasNext()) {
					Cell nextCell = cellIterator.next();

					int columnIndex = nextCell.getColumnIndex();

					switch (columnIndex) {
					case 0:
						String name = nextCell.getStringCellValue();
						statement.setString(1, name);
						break;
					case 1:
						String enrolled = nextCell.getStringCellValue();
						statement.setString(2, enrolled);
//					case 2:
//						int progress = (int) nextCell.getNumericCellValue();
//						statement.setInt(3, progress);
					}

				}
				
                statement.addBatch();
                
                if (count % batchSize == 0) {
                    statement.executeBatch();
                }				

			}

 
            workbook.close();
             
            // execute the remaining queries
            statement.executeBatch();
  
            connection.commit();
            connection.close();
             
            long end = System.currentTimeMillis();
            System.out.printf("Import done in %d ms\n", (end - start));
             
         }catch (IOException ex1) {
            System.out.println("Error reading file");
            ex1.printStackTrace();
        } catch (SQLException ex2) {
            System.out.println("Database error");
            ex2.printStackTrace();
        }
 
	}
}


	


