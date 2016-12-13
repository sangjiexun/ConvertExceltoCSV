package hoge;

import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelHandler {
	public static final String DOUBLEQUOT = "\"";

	public String getCellCSVValue(Workbook wb , Cell cell) {
		return getCellValue(wb , cell);
	}

	@SuppressWarnings("deprecation")
	public String getCellValue(Workbook wb , Cell cell) {
		String tempCellValue = "";

		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC: // 0
			if (DateUtil.isCellDateFormatted(cell)) {
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy/M/d");
				String strDate = sdf.format(cell.getDateCellValue()) + " 23:00";
//				pw.print(DOUBLEQUOT + strDate + DOUBLEQUOT);
				tempCellValue = DOUBLEQUOT + strDate + DOUBLEQUOT;
			} else {
//				pw.print(DOUBLEQUOT + cell.getNumericCellValue() + DOUBLEQUOT);
				tempCellValue = DOUBLEQUOT + cell.getNumericCellValue() + DOUBLEQUOT;
			}
			break;
		case Cell.CELL_TYPE_STRING: // 1
//			pw.print(DOUBLEQUOT + cell.getStringCellValue() + DOUBLEQUOT);
			tempCellValue = DOUBLEQUOT + cell.getStringCellValue() + DOUBLEQUOT;
			break;
		case Cell.CELL_TYPE_FORMULA: // 2
			CreationHelper crateHelper = wb.getCreationHelper();
			FormulaEvaluator evaluator = crateHelper.createFormulaEvaluator();
			evaluator.evaluateInCell(cell);
//			pw.print(DOUBLEQUOT + evaluator.evaluateInCell(cell) + DOUBLEQUOT);
			tempCellValue = DOUBLEQUOT + evaluator.evaluateInCell(cell) + DOUBLEQUOT;
			break;
		case Cell.CELL_TYPE_BLANK: // 3
//			pw.print(DOUBLEQUOT + "" + DOUBLEQUOT);
			tempCellValue = DOUBLEQUOT + "" + DOUBLEQUOT;
			break;
		case Cell.CELL_TYPE_BOOLEAN: // 4
//			pw.print(DOUBLEQUOT + cell.getBooleanCellValue() + DOUBLEQUOT);
			tempCellValue = DOUBLEQUOT + cell.getBooleanCellValue() + DOUBLEQUOT;
			break;
		case Cell.CELL_TYPE_ERROR: // 5
//			pw.print(DOUBLEQUOT + cell.getErrorCellValue() + DOUBLEQUOT);
			tempCellValue = DOUBLEQUOT + cell.getErrorCellValue() + DOUBLEQUOT;
			break;
		default:
			tempCellValue = DOUBLEQUOT + "" + DOUBLEQUOT;
		}
		return tempCellValue;
	}
}
