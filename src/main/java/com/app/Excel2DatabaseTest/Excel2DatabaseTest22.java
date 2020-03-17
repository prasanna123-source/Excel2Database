package com.app.Excel2DatabaseTest;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel2DatabaseTest22 {

	public static void main(String[] args) throws InvalidFormatException {
		// TODO Auto-generated method stub
		String jdbcURL = "jdbc:mysql://localhost:3306/employee";
		String username = "root";
		String password = "admin";

		// String excelFilePath = "Employee_data.xls";

		int batchSize = 50;

		Connection connection = null;

		try {
			long start = System.currentTimeMillis();

			// FileInputStream inputStream = new
			// FileInputStream("D:\\prasanna_workspace\\demo\\src\\main\\java\\com\\app\\Excel2DatabaseTest\\Employee_data.xls");
			Workbook wb = WorkbookFactory.create(new File(
					"D:\\prasanna_workspace\\demo\\src\\main\\java\\com\\app\\Excel2DatabaseTest\\Employee_data.xls"));
			Sheet mySheet = wb.getSheetAt(0);
			Iterator<Row> rowIterator = mySheet.rowIterator();
//            
//            Workbook workbook = new XSSFWorkbook(inputStream); 
//            org.apache.poi.ss.usermodel.Sheet firstSheet = workbook.getSheetAt(0);
//            Iterator<Row> rowIterator = firstSheet.iterator();

			connection = DriverManager.getConnection(jdbcURL, username, password);
			System.out.println("connection created");
			connection.setAutoCommit(false);

			String sql = "INSERT INTO employeedata (empName,bdate,salary) VALUES (?,?,?)";
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
//                    case 0:
//                    	int empId = (int)nextCell.getNumericCellValue();
//                        statement.setInt(1, empId);
//                        break;      
//                    
					case 0:
						String empName = nextCell.getStringCellValue();
						statement.setString(1, empName);
						break;
					case 1:
						Date bdate = nextCell.getDateCellValue();
						statement.setDate(2, new java.sql.Date(bdate.getTime()));
						break;
					case 2:
						int salary = (int) nextCell.getNumericCellValue();
						statement.setInt(3, salary);
						break;
					}

				}
				statement.execute();
				/*
				 * if (count % batchSize == 0) { statement.executeBatch(); }
				 */
			}

			wb.close();

			// execute the remaining queries
			// statement.executeBatch();

			connection.commit();
			connection.close();

			long end = System.currentTimeMillis();
			System.out.printf("Import done in %d ms\n", (end - start));

		} catch (IOException ex1) {
			System.out.println("Error reading file");
			ex1.printStackTrace();
		} catch (SQLException ex2) {
			System.out.println("Database error");
			ex2.printStackTrace();
		}

	}
}
