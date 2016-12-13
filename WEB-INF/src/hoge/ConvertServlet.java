package hoge;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConvertServlet extends HttpServlet {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public static final String COMMA = ",";
	public static final String DOUBLEQUOT = "\"";

	@SuppressWarnings("deprecation")
	public void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		System.out.println("START");

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

		// 出力先csvファイルを作成
		// 2番目の引数をtrueにすると追記モード、falseにすると上書きモード
		FileWriter fw = new FileWriter("/Users/aa352872/Desktop/JIKKEN.csv", false);
		PrintWriter pw = new PrintWriter(new BufferedWriter(fw));

		Sheet sheet = wb.getSheetAt(0);
		Row rowTemp = sheet.getRow(0);
		int lastRow = sheet.getLastRowNum();
		int lastCol = rowTemp.getLastCellNum();

		// System.out.println("行の最初行数" + sheet.getFirstRowNum());
		System.out.println("行の最大行数" + sheet.getLastRowNum());
		// System.out.println("行の最初列数" + rowTemp.getFirstCellNum());
		// これはなぜか1足された値になる(Java Docより)
		System.out.println("行の最大列数" + rowTemp.getLastCellNum());

		for (int rowNum = 0; rowNum <= lastRow; rowNum++) {
			Row row = sheet.getRow(rowNum);
			// System.out.println("getRow" + sheet.getRow(rowNum).toString());
			if (row != null) {
				for (int colNum = 0; colNum < lastCol; colNum++) {
					Cell cell = row.getCell(colNum);
					// System.out.println("getCell" + row.getCell(colNum));
					if (cell != null) {
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_NUMERIC: // 0
							if (DateUtil.isCellDateFormatted(cell)) {
								System.out.println("(row,col) = (" + rowNum + "," + colNum + ") " + "cell type:"
										+ cell.getCellType() + " " + "Date:" + cell.getDateCellValue());
								
								//SimpleDateFormat sdf = new SimpleDateFormat("yyyy/M/d HH:mm");
								SimpleDateFormat sdf = new SimpleDateFormat("yyyy/M/d");
						        //System.out.println(sdf.format(cell.getDateCellValue()));
								String strDate = sdf.format(cell.getDateCellValue()) + " 23:00";
								//pw.print(DOUBLEQUOT + cell.getDateCellValue() + DOUBLEQUOT);
								pw.print(DOUBLEQUOT + strDate + DOUBLEQUOT);
							} else {
								System.out.println("(row,col) = (" + rowNum + "," + colNum + ") " + "cell type:"
										+ cell.getCellType() + " " + "Numeric:" + cell.getNumericCellValue());
								pw.print(DOUBLEQUOT + cell.getNumericCellValue() + DOUBLEQUOT);
							}
							break;
						case Cell.CELL_TYPE_STRING: // 1
							System.out.println("(row,col) = (" + rowNum + "," + colNum + ") " + "cell type:"
									+ cell.getCellType() + " " + "String:" + cell.getStringCellValue());
							pw.print(DOUBLEQUOT + cell.getStringCellValue() + DOUBLEQUOT);
							break;
						case Cell.CELL_TYPE_FORMULA: // 2
							// System.out.println("Formula:" +
							// cell.getCellFormula());

							CreationHelper crateHelper = wb.getCreationHelper();
							FormulaEvaluator evaluator = crateHelper.createFormulaEvaluator();
							evaluator.evaluateInCell(cell);
							System.out.println("(row,col) = (" + rowNum + "," + colNum + ") " + "cell type:"
									+ cell.getCellType() + " " + "Formula:" + evaluator.evaluateInCell(cell));
							pw.print(DOUBLEQUOT + evaluator.evaluateInCell(cell) + DOUBLEQUOT);
							// pw.print(DOUBLEQUOT + cell.getCellFormula() +
							// DOUBLEQUOT);
							break;
						case Cell.CELL_TYPE_BLANK: // 3
							System.out.println("(row,col) = (" + rowNum + "," + colNum + ") " + "cell type:"
									+ cell.getCellType() + " " + "Blank:");
							pw.print(DOUBLEQUOT + "" + DOUBLEQUOT);
							break;
						case Cell.CELL_TYPE_BOOLEAN: // 4
							System.out.println("(row,col) = (" + rowNum + "," + colNum + ") " + "cell type:"
									+ cell.getCellType() + " " + "Boolean:" + cell.getBooleanCellValue());
							pw.print(DOUBLEQUOT + cell.getBooleanCellValue() + DOUBLEQUOT);
							break;
						case Cell.CELL_TYPE_ERROR: // 5
							System.out.println("(row,col) = (" + rowNum + "," + colNum + ") " + "cell type:"
									+ cell.getCellType() + " " + "Error:" + cell.getErrorCellValue());
							pw.print(DOUBLEQUOT + cell.getErrorCellValue() + DOUBLEQUOT);
							break;
						default:
						}
						// pw.print(DOUBLEQUOT + cell.getStringCellValue() +
						// DOUBLEQUOT);
						if (colNum != lastCol - 1) {
							pw.print(COMMA);
						}
					}
				}
			}
			pw.println();
		}
		// ファイルに書き出す
		pw.close();

		// workbookへ上書き
		FileOutputStream out = new FileOutputStream("/Users/aa352872/Desktop/sample.xlsx");
		wb.write(out);

		// Sheet sheet = wb.createSheet("new sheet");
		// //workbookを別名で保存
		// FileOutputStream out = null;
		// try{
		// out = new FileOutputStream("/Users/aa352872/Desktop/sample4_1.xlsx");
		// wb.write(out);
		// }catch(IOException e){
		// System.out.println(e.toString());
		// }finally{
		// try {
		// out.close();
		// }catch(IOException e){
		// System.out.println(e.toString());
		// }
		// }

		System.out.println("END\n-------");
		RequestDispatcher dispatcher = request.getRequestDispatcher("/jsp/Result.jsp");
		dispatcher.forward(request, response);
	}

}
