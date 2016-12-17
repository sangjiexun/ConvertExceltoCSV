package hoge;

import java.text.SimpleDateFormat;
import java.util.Iterator;

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
			sheetName = sheet.getSheetName();
			logger.trace("シート名 = " + sheetName);

			if (sheetName != "更新履歴" && sheetName != "記入例 " && sheetName != "各列の書き方 " && sheetName != "Ta運用フロー "
					&& sheetName != "各ケースの更新方法 " && sheetName != "テストケース作成更新した場合 ") {
				break;
			}
		}
		return sheet;
//		return "aaa";
//		return DOUBLEQUOT + getCellValue(wb , cell) + DOUBLEQUOT;
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
			//HACK
			//日付型 or 書式「全般」の和暦で20160101以降の日付
			if (DateUtil.isCellDateFormatted(cell) || cell.getNumericCellValue() > 42370) {
				strCellDateValue = sdf.format(cell.getDateCellValue()) + " 22:00";
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
