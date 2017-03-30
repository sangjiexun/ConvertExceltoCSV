package hoge;

import java.io.File;
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
		outputCSVFile = dirPath + "/" + OUTPUTFILENAME + sdfFileName.format(date).toString() + EXTENTION_CSV;
		logger.trace("出力CSVファイル = " + outputCSVFile);
		
		return outputCSVFile;
	}
	
	public String createOutputFolder(String uploadFolderPath){
		Date date = new Date();
        SimpleDateFormat sdfFileName = new SimpleDateFormat("yyyyMMddHHmmss");
        File o = new File(uploadFolderPath + "/csv" + sdfFileName.format(date).toString());
        o.mkdir();
        
		return o.toString();
	}
	
}
