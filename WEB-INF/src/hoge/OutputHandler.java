package hoge;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.log4j.Logger;

public class OutputHandler {
	final Logger logger = Logger.getLogger (OutputHandler.class);
	public static final String OUTPUTFILENAME = "forImport";
	public static final String EXTENTION_CSV = ".csv";
	
	public String createOutputFile(String dirPath){
		String outputCSVFile = "";
		
		Date date = new Date();
        SimpleDateFormat sdfFileName = new SimpleDateFormat("yyyyMMddHHmmss");
		outputCSVFile = dirPath + OUTPUTFILENAME + sdfFileName.format(date).toString() + EXTENTION_CSV;
		logger.trace("出力CSVファイル = " + outputCSVFile);
		
		return outputCSVFile;
	}
}
