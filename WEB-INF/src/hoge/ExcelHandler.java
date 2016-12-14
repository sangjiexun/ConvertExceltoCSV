package hoge;

import java.text.SimpleDateFormat;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;


public class ExcelHandler {
	final Logger logger = Logger.getLogger(ExcelHandler.class);
	public static final String DOUBLEQUOT = "\"";
	
	//cellの内容を取得する
	public String getCellCSVValue(Workbook wb , Cell cell) {
		return getCellValue(wb , cell);
	}
	
	//TYPEに合わせてcellの内容を取得する
	@SuppressWarnings("deprecation")
	public String getCellValue(Workbook wb , Cell cell) {
		String tempCellValue = "";

		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC: // 0
			if (DateUtil.isCellDateFormatted(cell)) {
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy/M/d");
				String strDate = sdf.format(cell.getDateCellValue()) + " 22:00";
				logger.trace("Cell.CELL_TYPE_NUMERIC:" + "(row.col) = (" + cell.getRowIndex() + "," + cell.getColumnIndex() + ")" + "cellValue = " + strDate);
				tempCellValue = DOUBLEQUOT + strDate + DOUBLEQUOT;
			} else {
				//TODO
				//HACK
				if(cell.getNumericCellValue() > 42370){
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy/M/d");
					String strDate = sdf.format(cell.getDateCellValue()) + " 22:00";
					logger.trace("DDDD = " + strDate);
				}
				
				logger.trace("Cell.CELL_TYPE_NUMERIC:" + "(row.col) = (" + cell.getRowIndex() + "," + cell.getColumnIndex() + ")" + "cellValue = " + cell);
				tempCellValue = DOUBLEQUOT + cell.getNumericCellValue() + DOUBLEQUOT;
			}
			break;
		case Cell.CELL_TYPE_STRING: // 1
			logger.trace("Cell.CELL_TYPE_STRING:" + "(row.col) = (" + cell.getRowIndex() + "," + cell.getColumnIndex() + ")" + "cellValue = " + cell);
			tempCellValue = DOUBLEQUOT + cell.getStringCellValue() + DOUBLEQUOT;
			break;
		case Cell.CELL_TYPE_FORMULA: // 2
			CreationHelper crateHelper = wb.getCreationHelper();
			FormulaEvaluator evaluator = crateHelper.createFormulaEvaluator();
			evaluator.evaluateInCell(cell);
			logger.trace("Cell.CELL_TYPE_FORMULA:" + "(row.col) = (" + cell.getRowIndex() + "," + cell.getColumnIndex() + ")" + "cellValue = " + cell);
			tempCellValue = DOUBLEQUOT + evaluator.evaluateInCell(cell) + DOUBLEQUOT;
			break;
		case Cell.CELL_TYPE_BLANK: // 3
			logger.trace("Cell.CELL_TYPE_BLANK:" + "(row.col) = (" + cell.getRowIndex() + "," + cell.getColumnIndex() + ")" + "cellValue = " + cell);
			tempCellValue = DOUBLEQUOT + "" + DOUBLEQUOT;
			break;
		case Cell.CELL_TYPE_BOOLEAN: // 4
			logger.trace("Cell.CELL_TYPE_BOOLEAN:" + "(row.col) = (" + cell.getRowIndex() + "," + cell.getColumnIndex() + ")" + "cellValue = " + cell);
			tempCellValue = DOUBLEQUOT + cell.getBooleanCellValue() + DOUBLEQUOT;
			break;
		case Cell.CELL_TYPE_ERROR: // 5
			logger.trace("Cell.CELL_TYPE_ERROR:" + "(row.col) = (" + cell.getRowIndex() + "," + cell.getColumnIndex() + ")" + "cellValue = " + cell);
			tempCellValue = DOUBLEQUOT + cell.getErrorCellValue() + DOUBLEQUOT;
			break;
		default:
			logger.trace("default:" + "(row.col) = (" + cell.getRowIndex() + "," + cell.getColumnIndex() + ")" + "cellValue = " + cell);
			tempCellValue = DOUBLEQUOT + "" + DOUBLEQUOT;
		}
		return tempCellValue;
	}
}
