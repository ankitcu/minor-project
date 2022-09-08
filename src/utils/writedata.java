package utils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writedata {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
   HSSFWorkbook workbook= new HSSFWorkbook();
    HSSFSheet sheet1 = workbook.createSheet("Sheet1");
    
     Row r0= sheet1.createRow(0);
    Cell c0= r0.createCell(0);
    c0.setCellValue("Subject no..");
    
    Cell c01= r0.createCell(1);
    c01.setCellValue("Grades");
    
    Cell c02= r0.createCell(2);
    c02.setCellValue("Semester");
    
    Row r1= sheet1.createRow(1);
    Cell c11= r1.createCell(0);
    c11.setCellValue("1");
   
    Cell c12= r1.createCell(1);
    c12.setCellValue("O");
    
    Cell c13= r1.createCell(2);
    c13.setCellValue("4");
    
    Row r2= sheet1.createRow(2);
    Cell c21= r2.createCell(0);
    c21.setCellValue("2");
   
    Cell c22= r2.createCell(1);
    c22.setCellValue("A+");
    
    Cell c23= r2.createCell(2);
    c23.setCellValue("4");
    
    Row r3= sheet1.createRow(3);
    Cell c31= r3.createCell(0);
    c31.setCellValue("3");
   
    Cell c32= r3.createCell(1);
    c32.setCellValue("A+");
   
    Cell c33= r3.createCell(2);
    c33.setCellValue("4");
    
    Row r4= sheet1.createRow(4);
    Cell c41= r4.createCell(0);
    c41.setCellValue("4");
    
    Cell c42= r4.createCell(1);
    c42.setCellValue("A+");
   
    Cell c43= r4.createCell(2);
    c43.setCellValue("4");
    
    Row r5= sheet1.createRow(5);
    Cell c51= r5.createCell(0);
    c51.setCellValue("5");
    
    Cell c52= r5.createCell(1);
    c52.setCellValue("O");
   
    Cell c53= r5.createCell(2);
    c53.setCellValue("4");
   
    Row r6= sheet1.createRow(6);
    Cell c61= r6.createCell(0);
    c61.setCellValue("6");
    
    Cell c62= r6.createCell(1);
    c62.setCellValue("O");
   
    Cell c63= r6.createCell(2);
    c63.setCellValue("4");
    
    Row r7= sheet1.createRow(7);
    Cell c71= r7.createCell(0);
    c71.setCellValue("7");
    
    Cell c72= r7.createCell(1);
    c72.setCellValue("O");
   
    Cell c73= r7.createCell(2);
    c73.setCellValue("4");
    
    Row r8= sheet1.createRow(8);
    Cell c81= r8.createCell(0);
    c81.setCellValue("8");
    
    Cell c82= r8.createCell(1);
    c82.setCellValue("A");
   
    Cell c83= r8.createCell(2);
    c83.setCellValue("4");
    
    Row r9= sheet1.createRow(9);
    Cell c91= r9.createCell(0);
    c91.setCellValue("9");
    
    Cell c92= r9.createCell(1);
    c92.setCellValue("A+");
   
    Cell c93= r9.createCell(2);
    c93.setCellValue("4");
    
    Row r10= sheet1.createRow(10);
    Cell c101= r10.createCell(0);
    c101.setCellValue("10");
    
    Cell c102= r10.createCell(1); 
    c102.setCellValue("A+");
   
    Cell c103= r10.createCell(2);
    c103.setCellValue("4");
    File f=new File("C:\\Users\\abhishek\\Desktop\\java\\learning\\Minorproject\\src\\utils\\writedata.xls");
    FileOutputStream fos=  new FileOutputStream(f);
    workbook.write(fos);
    fos.close();
    workbook.close();
    System.out.println("File is Written Successfully!");
	}

}