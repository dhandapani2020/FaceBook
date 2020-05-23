import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class LoginPage {

	public static void main(String[] args) throws IOException {
				
		/*//Read Excel
		File f=new File("C:\\Users\\dhand\\Desktop\\Test.xlsx");
		FileInputStream fis=new FileInputStream(f);
		Workbook w=new XSSFWorkbook(fis);
		Sheet s=w.getSheet("Sheet1");
		Row r=s.getRow(0);
		Cell c=r.getCell(0);
		System.out.println(c);
		
		
		//Excel Write		
		Workbook w1=new XSSFWorkbook(fis);
		Sheet s1=w1.getSheet("Sheet1");
		Row r1=s1.createRow(5);
		Cell C1=r1.createCell(4);
		C1.setCellValue("Java");
		FileOutputStream fos=new FileOutputStream(f);
		w1.write(fos);
		System.out.println("Done");*/
		
		/*//Create new Excel file
		File f1=new File("C:\\Users\\dhand\\Desktop\\Test2.xlsx");
		FileInputStream fis=new FileInputStream(f1);
		Workbook w2=new XSSFWorkbook(fis);
		Sheet s2=w2.createSheet("Abi");
		System.out.println("File created");
		Row r1=s2.createRow(0);
		Cell c1 = r1.createCell(0);
		c1.setCellValue("Java");
		System.out.println("Sheet Updated");
		*/
		
		//To create a new WorkBook with xlsx extension		
		 
		FileOutputStream fileOut1 = new FileOutputStream("D:\\Test9.xls");
		Workbook wb1 = new XSSFWorkbook();		
		
		Sheet s2=wb1.createSheet("Abi");
		System.out.println("File created");
		Row r1=s2.createRow(0);
		Cell c1 = r1.createCell(0);
		c1.setCellValue("Java");
		wb1.write(fileOut1);
		}
}

