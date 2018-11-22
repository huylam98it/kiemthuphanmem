package kiemthu;

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

public class DatBan {
	static WebDriver webDriver;

	public static void main(String[] args) throws IOException {
		System.setProperty("webdriver.gecko.driver", "src/geckodriver");
		webDriver = new FirefoxDriver();
		
		loadData("C:\\Users\\huylam98it\\Documents\\kiemthu\\Giang_TC.xlsx", 3, 1, 5);
	}

	public static void loadData(String path, int firstRow, int firstCell, int lastCell) throws IOException {
		FileInputStream inputStream = new FileInputStream(new File(path));
		XSSFWorkbook workBook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workBook.getSheetAt(1);
		int lastRow = sheet.getLastRowNum();
		System.out.println("firstRow: " + firstRow);
		System.out.println("lastRow: " + lastRow);
		System.out.println("firstCell: " + firstCell);
		System.out.println("lastCell: " + lastCell);

		for (int i = firstRow; i <= lastRow; i++) {
			Row row = sheet.getRow(i);
			ArrayList<String> data = new ArrayList<>();
			for (int j = firstCell; j <= lastCell; j++) {
				data.add(getCellData(row.getCell(j)));
			}
			for (String a : data) {
				System.out.print(a + "+");
			}
			String kq = row.getCell(lastCell + 1).getStringCellValue().toString().trim();

			boolean result = kq.equals(mainTest(data));

			Cell round = row.createCell(row.getLastCellNum());
			round.setCellValue(result);
		}
		inputStream.close();
		FileOutputStream outputStream = new FileOutputStream(new File(path));
		workBook.write(outputStream);
		outputStream.close();
		workBook.close();
	}

	public static String mainTest(ArrayList<String> data) {
		webDriver.get("http://localhost:8080/BTLnhom16/datbanonline");

		WebElement ten=webDriver.findElement(By.name("name"));
		WebElement address=webDriver.findElement(By.name("address"));
		Select tg=new Select(webDriver.findElement(By.name("thoigian")));
		WebElement soluong=webDriver.findElement(By.name("soluong"));
		WebElement phone=webDriver.findElement(By.name("dienthoai"));
		
		if(data.get(0)!=null) {
			ten.sendKeys(data.get(0));
		}
		if(data.get(1)!=null) {
			address.sendKeys(data.get(1));
		}
		if(data.get(2)!=null) {
			tg.selectByVisibleText(data.get(2));
		}
		if(data.get(3)!=null) {
			soluong.sendKeys(data.get(3));
		}
		if(data.get(4)!=null) {
			phone.sendKeys(data.get(4));
		}
		
		webDriver.findElement(By.id("btn-add")).click();
		
		return webDriver.findElement(By.id("kq")).getText();
	}

	public static String getCellData(Cell cell) {
		String kq = null;
		try {
			cell.getCellTypeEnum();
		} catch (Exception e) {

			return kq;
		}
		switch (cell.getCellTypeEnum()) {
		case STRING:
			kq = cell.getStringCellValue().toString().trim();
			break;
		case NUMERIC:
			int tam = (int) (cell.getNumericCellValue());
			kq = tam + "";
		default:

			break;
		}
		return kq;
	}
}
