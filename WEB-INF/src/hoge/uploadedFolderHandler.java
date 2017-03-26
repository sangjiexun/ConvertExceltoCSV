package hoge;

import java.io.File;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class uploadedFolderHandler {
	
	private static final String UPLOADEDFOLDER = "uploadedFolder";
	
	public String createUploadedFolder(String realPath) {
		//uploadファイル保存用フォルダ作成
		File uploadedFolder = new File(realPath + UPLOADEDFOLDER);
		if(!uploadedFolder.exists()){
			uploadedFolder.mkdir();
			//System.out.println("uploadedFolder = " + uploadedFolder);
		}
		return uploadedFolder.toString();
	}
	
	public String createTargetFolder(String uploadedFolder){
		LocalDateTime dt = LocalDateTime.now();
		//System.out.println("local time = " + d.toString());
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
		String uploadedFolderName = dt.format(formatter);
		//System.out.println("hiduke = " + uploadedFolderName);
		File targetFolder = new File(uploadedFolder + "/" + uploadedFolderName);
		targetFolder.mkdir();
		//System.out.println("timeStampFolder = " + targetFolder);
		
		return targetFolder.toString();
	}
	

}
