package hoge;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelHandler {

	public Workbook openExistWorkbook(String targetDir, String targetFile) {

		FileInputStream in = null;
		Workbook wb = null;

		// 既存のworkbookを開く
		try {
			in = new FileInputStream("/Users/aa352872/Desktop/sample.xlsx");
			wb = WorkbookFactory.create(in);
		} catch (IOException e) {
			System.out.println(e.toString());
		} catch (InvalidFormatException e) {
			System.out.println(e.toString());
		} finally {
			try {
				in.close();
			} catch (IOException e) {
				System.out.println(e.toString());
			}
		}
		return wb;
	}

}
