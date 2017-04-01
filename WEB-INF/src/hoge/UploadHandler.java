package hoge;

import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.Part;

public class UploadHandler {
	
	private static final String UPLOADEDFOLDER = "uploadedFolder";
	
	//全部loggerに変更
	
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
	
	public void writeUploadedFile(HttpServletRequest request,String targetFolder) throws IOException, ServletException{
		//アップロードファイル全てを取得する
		for (Part part : request.getParts()) {
			//System.out.println("parts = " + part);
			String uploadedFileName = this.getFileName(part);
			//System.out.println("uploadedFileName = " + uploadedFileName);
			//アップロードファイルを指定フォルダに書き出す
			part.write(targetFolder + "/" + uploadedFileName);
		}
	}
	
	private String getFileName(Part part) {
        String name = null;
        for (String dispotion : part.getHeader("Content-Disposition").split(";")) {
            if (dispotion.trim().startsWith("filename")) {
                name = dispotion.substring(dispotion.indexOf("=") + 1).replace("\"", "").trim();
                name = name.substring(name.lastIndexOf("\\") + 1);
                break;
            }
        }
        return name;
    }
	

}
