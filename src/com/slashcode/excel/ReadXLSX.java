package com.slashcode.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.slashcode.pojo.Customer;

public class ReadXLSX {

	public static void main(String[] args) {
		List<Customer> custList = readXLSXFile("D:\\Test\\CustDetailsXLSX.xlsx");
		for(Customer cust : custList){
			System.out.println(cust);
		}
	}

	private static List<Customer> readXLSXFile(String file) {
		List<Customer> listCust = new ArrayList<Customer>();
		try {
			XSSFWorkbook work = new XSSFWorkbook(new FileInputStream(file));
			
			XSSFSheet sheet = work.getSheet("Customer");
			XSSFRow row = null;
			
			int i=0;
			while((row = sheet.getRow(i))!=null){
				int custId,pinCode;
				String custName,custCity,stateCode;
				try{
					custId = (int) row.getCell(0).getNumericCellValue();
				}
				catch(Exception e){custId = 0;}
				try{
					custName = row.getCell(1).getStringCellValue();
				}
				catch(Exception e){custName = null;}
				try{
					custCity = row.getCell(2).getStringCellValue();
				}
				catch(Exception e){custCity = null;}
				try{
					pinCode = (int) row.getCell(3).getNumericCellValue();
				}
				catch(Exception e){pinCode = 0;}
				try{
					stateCode = row.getCell(4).getStringCellValue();
				}
				catch(Exception e){stateCode = null;}
				Customer cust = new Customer(custId,custName,custCity,pinCode,stateCode);
				listCust.add(cust);
					i++;				
			}
			work.close();
		} catch (IOException e) {
			System.out.println("Exception is Customer fetch data :: "+e.getMessage());
			e.printStackTrace();
		}
		return listCust;
	}	
}

