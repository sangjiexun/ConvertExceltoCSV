package hoge;

import java.text.SimpleDateFormat;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class ExcelHandler {
	final Logger logger = Logger.getLogger(ExcelHandler.class);
	public static final String DOUBLEQUOT = "\"";
	
	//sheet名を取得する
	public Sheet getSheetName(Workbook wb) {
	 	String sheetName = "";
	 	Sheet sheet = null;

	 	Iterator<Sheet> sheets = wb.sheetIterator();
	 	
		while (sheets.hasNext()) {
			sheet = sheets.next();
			sheetName = sheet.getSheetName().trim();

			if (!sheetName.equals("更新履歴") && !sheetName.equals("記入例") && !sheetName.equals("各列の書き方")
					&& !sheetName.equals("Ta運用フロー") && !sheetName.equals("各ケースの更新方法")
					&& !sheetName.equals("テストケース作成更新した場合")) {
				logger.trace("HITシート名 = " + sheetName);
				return sheet;
			}
		}
		return sheet;
	}
	
	//cellの内容を取得する
	public String getCellCSVValue(Workbook wb , Cell cell) {
		return DOUBLEQUOT + getCellValue(wb , cell) + DOUBLEQUOT;
	}
	
	//TYPEに合わせてcellの内容を取得する
	@SuppressWarnings("deprecation")
	public String getCellValue(Workbook wb , Cell cell) {
		String strCellIndex = "(row.col) = (" + cell.getRowIndex() + "," + cell.getColumnIndex() + ") ";
		String tempCellValue = "";
		String strCellDateValue = "";
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy/M/d");

		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC: // 0
			//日付タイプと数値タイプは同じくCELL_TYPE_NUMERIC
			//数値か日付かを判別するためにはDateUtilクラスで用意されているisCellDateFormattedメソッドを使用して判断する
			//isCellDateFormattedがtrueを返した場合はセルのタイプは日付、falseを返した場合は数値タイプ
			if (DateUtil.isCellDateFormatted(cell)) {
				strCellDateValue = sdf.format(cell.getDateCellValue());
				logger.trace("Cell.CELL_TYPE_NUMERIC:" + strCellIndex + "cellValue = " + strCellDateValue);
				tempCellValue = strCellDateValue;
			} else {
				logger.trace("Cell.CELL_TYPE_NUMERIC:" + strCellIndex + "cellValue = " + cell);
				tempCellValue = String.valueOf(cell.getNumericCellValue());
			}
			break;
		case Cell.CELL_TYPE_STRING: // 1
			logger.trace("Cell.CELL_TYPE_STRING:" + strCellIndex + "cellValue = " + cell);
			tempCellValue = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_FORMULA: // 2
			//セルの値の計算結果を取得する
			CreationHelper crateHelper = wb.getCreationHelper();
			FormulaEvaluator evaluator = crateHelper.createFormulaEvaluator();
			evaluator.evaluateInCell(cell);
			logger.trace("Cell.CELL_TYPE_FORMULA:" + strCellIndex + "cellValue = " + cell);
			tempCellValue = String.valueOf(evaluator.evaluateInCell(cell));
			break;
		case Cell.CELL_TYPE_BLANK: // 3
			logger.trace("Cell.CELL_TYPE_BLANK:" + strCellIndex + "cellValue = " + cell);
			tempCellValue = "";
			break;
		case Cell.CELL_TYPE_BOOLEAN: // 4
			logger.trace("Cell.CELL_TYPE_BOOLEAN:" + strCellIndex + "cellValue = " + cell);
			tempCellValue = String.valueOf(cell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_ERROR: // 5
			logger.trace("Cell.CELL_TYPE_ERROR:" + strCellIndex + "cellValue = " + cell);
			tempCellValue = String.valueOf(cell.getErrorCellValue());
			break;
		default:
			logger.trace("default:" + strCellIndex + "cellValue = " + cell);
			tempCellValue = "";
		}
		return tempCellValue;
	}
	
	
}
