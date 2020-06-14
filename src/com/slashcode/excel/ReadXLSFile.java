package com.slashcode.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.slashcode.pojo.Customer;

public class ReadXLSFile {

	public static void main(String[] args) {
		try {
			List<Customer> custList = readXLSFile("D:\\Test\\CustDetails.xls");
			for(Customer cust : custList){
				System.out.println(cust);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}		
	}

	private static List<Customer> readXLSFile(String file) throws FileNotFoundException, IOException {
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
		HSSFSheet sheet = workbook.getSheet("Customer");
		List<Customer> listCust = new ArrayList<Customer>();
		HSSFRow row = null;
		int i=0;
		while((row=sheet.getRow(i)) != null){
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
		workbook.close();
		return listCust;
	}
}
