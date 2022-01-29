package application;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		new Main().exe();
		System.out.println("Fim.");
	}

	@SuppressWarnings("resource")
	public void exe() {
		String local = "C:\\Temp";
		String nome = "arq.xlsx";
		File file = new File(local);
		if(!file.exists())
			file.mkdir();
		Workbook wb = new XSSFWorkbook();
		Sheet sh = wb.createSheet("okOk");
		sh.createRow(0).createCell(0).setCellValue("50");
		try (FileOutputStream fos = new FileOutputStream(new File(local + File.separator + nome))) {
			wb.write(fos);
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
