package hoge;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ConvertServlet extends HttpServlet {
	final Logger logger = Logger.getLogger(ConvertServlet.class);
	private static final long serialVersionUID = 1L;

	public static final String COMMA = ",";
	// public static final String DOUBLEQUOT = "\"";
	public String TARGETDIR = "";

	public void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		logger.trace("START CONVERT PROCESS");
		// nullString Validation check
		Validation vali = new Validation();
		if (!vali.checkNullString(request.getParameter("targetDir"))) {
			request.setAttribute("InputError", "nullString");
			RequestDispatcher dispatcher = request.getRequestDispatcher("/jsp/InputError.jsp");
			dispatcher.forward(request, response);
			return;
		} else {
			TARGETDIR = request.getParameter("targetDir");
		}

		// 出力CSVファイル生成
		OutputFileHandler ofh = new OutputFileHandler();
		String outputCSVFile = ofh.createOutputFile("/Users/aa352872/Desktop/");

		File dir = new File(TARGETDIR);
		String[] files = dir.list();
		FileInputStream in = null;
		Workbook wb = null;

		logger.trace("選択フォルダ内のファイル合計 = " + files.length);
		int convertExcelFiles = 0;
		for (int i = 0; i < files.length; i++) {
			// 既存のworkbookを開く
			try {
				// エクセルファイルのみを対象に読み込む
				if (files[i].indexOf("xls") != -1) {
					convertExcelFiles++;
					logger.trace("convert対象ファイル = " + files[i] + " 読み込みファイル数 = " + convertExcelFiles);
					in = new FileInputStream(TARGETDIR + "/" + files[i]);
				} else {
					logger.trace("skipファイル = " + files[i]);
					continue;
				}
				wb = WorkbookFactory.create(in);
			} catch (IOException e) {
				System.out.println(e.toString());
			} catch (InvalidFormatException e) {
				System.out.println(e.toString());
			} finally {
				try {
					if (in != null) {
						in.close();
					}
				} catch (IOException e) {
					System.out.println(e.toString());
				}
			}

			// 出力先csvファイルを作成
			// 2番目の引数をtrueにすると追記モード、falseにすると上書きモード
			FileWriter fw = new FileWriter(outputCSVFile, true);
			PrintWriter pw = new PrintWriter(new BufferedWriter(fw));

			// 1シート目を取得
			Sheet sheet = wb.getSheetAt(0);
			// 1行目を取得
			Row rowTemp = sheet.getRow(0);
			// 行数を取得
			int lastRow = sheet.getLastRowNum();
			logger.trace("対象ファイル最大行数 = " + lastRow);
			// 列数を取得
			int lastCol = rowTemp.getLastCellNum(); // これはなぜか1足された値になる(JavaDocより)
			logger.trace("対象ファイル最大列数 + 1 = " + lastCol);

			ExcelHandler exh = new ExcelHandler();

			for (int rowNum = 0; rowNum <= lastRow; rowNum++) {
				// 2ファイルめ以降の1行目(ヘッダー)をskipする
				if (convertExcelFiles != 1 && rowNum == 0) {
					continue;
				}
				Row row = sheet.getRow(rowNum);
				
				if (row != null) {
					for (int colNum = 0; colNum < lastCol; colNum++) {
						Cell cell = row.getCell(colNum);
						//cellの内容を取得
						if (cell != null) {
							pw.print(exh.getCellCSVValue(wb, cell));
							if (colNum != lastCol - 1) {
								pw.print(COMMA);
							}
						}
					}
					//最後に改行を入れる
					pw.println();
				}else{
					//Nothing
				}
			}
			// ファイルに書き出す
			pw.close();
		}
		logger.trace("出力ファイル合計 = " + convertExcelFiles);
		logger.trace("END CONVERT PROCESS");

		RequestDispatcher dispatcher = request.getRequestDispatcher("/jsp/Result.jsp");
		dispatcher.forward(request, response);
	}
}
