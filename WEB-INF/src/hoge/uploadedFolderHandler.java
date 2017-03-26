package hoge;

import java.io.File;

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
	

}
