package PhamHuyLam;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

public class ThemMonAn {
	
	static WebDriver webDriver;
	
	public static void main(String[] args) throws IOException {
		System.setProperty("webdriver.gecko.driver","src/geckodriver.exe");
		webDriver=new FirefoxDriver();
		//dang nhap tai khoan admin
		webDriver.get("http://localhost:8080/BTLnhom16/login");
		WebElement username= webDriver.findElement(By.name("username"));
		WebElement password=webDriver.findElement(By.name("password"));
		username.sendKeys("admin");
		password.sendKeys("admin");
		webDriver.findElement(By.id("btn-login")).click();;
		
		loadData("src/Lam_datatest.xlsx", 2, 1, 5);
	}
	
	public static void loadData(String path,int firstRow,int firstCell,int lastCell) throws IOException {
		FileInputStream inputStream=new FileInputStream(new File(path));
		XSSFWorkbook workBook=new XSSFWorkbook(inputStream);
		XSSFSheet sheet=workBook.getSheetAt(0);
		int lastRow=sheet.getLastRowNum();
		System.out.println("firstRow: "+firstRow);
		System.out.println("lastRow: "+lastRow);
		System.out.println("firstCell: "+firstCell);
		System.out.println("lastCell: "+lastCell);
		int stt=1;
		
		for(int i=firstRow;i<=lastRow;i++) {
			System.out.println("test case "+stt+" dang chay");
			stt+=1;
			Row row=sheet.getRow(i);
			ArrayList<String> data=new ArrayList<>();
			for(int j=firstCell;j<=lastCell;j++) {
				data.add(getCellData(row.getCell(j)));
			}
			for(String a:data) {
				System.out.print(a+"+");
			}
			String kq=row.getCell(lastCell+1).getStringCellValue().toString().trim();
			
			boolean result=kq.equals(mainTest(data));
			Cell round=row.createCell(row.getLastCellNum());
			round.setCellValue(result);
		}
		inputStream.close();
		FileOutputStream outputStream=new FileOutputStream(new File(path));
		workBook.write(outputStream);
		outputStream.close();
		workBook.close();
	}
	
	public static String mainTest(ArrayList<String>data) {
		webDriver.get("http://localhost:8080/BTLnhom16/themmonan");//den trang them mon an
		
		WebElement ten=webDriver.findElement(By.name("namemonan"));
		Select loai=new Select(webDriver.findElement(By.name("idtheloai")));
		WebElement hinh=webDriver.findElement(By.name("filehinhanh"));
		Select thoigian=new Select(webDriver.findElement(By.name("thoigian")));
		WebElement gia=webDriver.findElement(By.name("giatien"));
		
		//thuc hien truyen du lieu
		if(data.get(0)!=null) {
			System.out.println("dang set ten");
			ten.sendKeys(data.get(0));
		}
		if(data.get(1)!=null) {
			loai.selectByVisibleText(data.get(1));
		}
		if(data.get(2)!=null) {
			hinh.sendKeys(data.get(2));
			try {
				Thread.sleep(1000);
				//webDriver.switchTo().alert().accept();
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
		if(data.get(3)!=null) {
			thoigian.selectByVisibleText(data.get(3));
		}
		if(data.get(4)!=null) {
			gia.sendKeys(data.get(4));
		}
		
		//gui du lieu
		webDriver.findElement(By.id("btn-add")).click();
		System.out.println("thong bao "+webDriver.findElement(By.id("tb")).getText());
		return webDriver.findElement(By.id("tb")).getText();
		
	}
	
	
	
	public static String getCellData(Cell cell) {
		String kq=null;
		try {
			cell.getCellTypeEnum();
		}
		catch (Exception e) {
			
			return kq;
		}
		switch (cell.getCellTypeEnum()) {
		case STRING:
			kq=cell.getStringCellValue().toString().trim();
			break;
		case NUMERIC:
			int tam=(int)(cell.getNumericCellValue());
			kq=tam+"";
		default:
			
			break;
		}
		return kq;
	}
	
}
