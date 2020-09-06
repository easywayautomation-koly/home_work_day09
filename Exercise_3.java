package home_work_day09;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import javax.imageio.stream.FileImageOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Exercise_3 {

	public static void main(String[] args) throws Exception {

		FileInputStream fis = new FileInputStream("C:\\Book1.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("Sheet1");
		// FormulaEvaluator formulaEvaluator
		// =wb.getCreationHelper().createFormulaEvaluator();

		for (Row row : sheet) {
			
			for (Cell Cell : row) {
				System.out.print(Cell+" ");
				
			}
			System.out.println();
		}

	}

}
