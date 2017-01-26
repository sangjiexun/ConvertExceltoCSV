package hoge;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;

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
		Sheet sheet = null;
		ArrayList<String> arrayInFile = new ArrayList<String>();
		ArrayList<String> arraySkipFile = new ArrayList<String>();

		logger.trace("選択フォルダ内のファイル合計 = " + files.length);
		int convertExcelFiles = 0;
		for (int i = 0; i < files.length; i++) {
			// 既存のworkbookを開く
			try {
				// エクセルファイルのみを対象に読み込む
				if (files[i].indexOf("xls") != -1) {
					convertExcelFiles++;
					arrayInFile.add(files[i]);
					logger.trace("convert対象ファイル = " + files[i] + " 読み込みファイル数 = " + convertExcelFiles);
					in = new FileInputStream(TARGETDIR + "/" + files[i]);
				} else {
					arraySkipFile.add(files[i]);
					logger.trace("skipファイル(Excelでない) = " + files[i]);
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

			ExcelHandler exh = new ExcelHandler();
			// 1シート目を取得
			// Sheet sheet = wb.getSheetAt(0);
			// シートを順番に調べて対象を返す
			sheet = exh.getSheetName(wb);
			if (sheet == null) {
				continue;
			}

			// 1行目を取得
			Row rowTemp = sheet.getRow(0);
			// 行数を取得
			int lastRow = sheet.getLastRowNum();
			logger.trace("対象ファイル最大行数 = " + lastRow);
			// 列数を取得
			int lastCol = rowTemp.getLastCellNum(); // これはなぜか1足された値になる(JavaDocより)
			logger.trace("対象ファイル最大列数 + 1 = " + lastCol);
			//状況列の列index
			int statusColIndex = 9999;
			//ID列の列index
			int IDColIndex = 9999;

			for (int rowNum = 0; rowNum <= lastRow; rowNum++) {
				// 2ファイルめ以降の1行目(ヘッダー)をskipする
				if (convertExcelFiles != 1 && rowNum == 0) {
					continue;
				}
				Row row = sheet.getRow(rowNum);

				if (row != null) {
					for (int colNum = 0; colNum < lastCol; colNum++) {

						Cell cell = row.getCell(colNum);
						// cellの内容を取得
						if (cell != null) {
							
							String outputCellValue = exh.getCellCSVValue(wb, cell);
							//HACK
							//状況列indexを記録する
							if(cell.getRowIndex() == 0 && outputCellValue.equals("\"状況\"")){
								statusColIndex = colNum;
							}
							
							
							// 1列目のcellがblankなら次の行へ
							// 1列目のcellがスペースのみなら次の行へ
							// HACK
							if (colNum == 0 && (outputCellValue.equals("\"\"") || outputCellValue.equals("\"　\"")
									|| outputCellValue.equals("\"____\""))) {
								break;
							}
							
							if(colNum == statusColIndex){
								if(outputCellValue.equals("\"テスト実施完了\"")){
									outputCellValue = "テスト結果検証";
								}else if(outputCellValue.equals("\"テスト結果検証完了\"")){
									outputCellValue = "完了";
								}
							}

							pw.print(outputCellValue);
							if (colNum != lastCol - 1) {
								pw.print(COMMA);
							} else {
								if(!request.getParameter("parentCategory").equals("nothing")){
									if(rowNum == 0){
										pw.print(COMMA + "テストケース親子区分");
									}else if(request.getParameter("parentCategory").equals("parent")){
										pw.print(COMMA + "\"親\"");
									}else if(request.getParameter("parentCategory").equals("child")){
										pw.print(COMMA + "\"子\"");
									}
								}
								if(!request.getParameter("group").equals("nothing")){
									if(rowNum == 0){
										pw.print(COMMA + "検出元");
									}else if(request.getParameter("group").equals("A")){
										pw.print(COMMA + "\"GroupA\"");
									}else if(request.getParameter("group").equals("B")){
										pw.print(COMMA + "\"GroupB\"");
									}else if(request.getParameter("group").equals("C")){
										pw.print(COMMA + "\"GroupC\"");
									}
								}
								
								// 最後に改行を入れる
								pw.println();
							}
						} else {
							if (colNum == 0) {
								break;
							}
						}
					}
				} else {
					// Nothing
				}
			}
			// ファイルに書き出す
			pw.close();
		}
		request.setAttribute("arrayInFile", arrayInFile);
		request.setAttribute("arraySkipFile", arraySkipFile);
		logger.trace("出力ファイル合計 = " + convertExcelFiles);
		logger.trace("END CONVERT PROCESS");

		RequestDispatcher dispatcher = request.getRequestDispatcher("/jsp/Result.jsp");
		dispatcher.forward(request, response);
	}
}
