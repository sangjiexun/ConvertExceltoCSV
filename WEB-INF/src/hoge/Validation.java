package hoge;

import java.io.IOException;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.Part;

import org.apache.log4j.Logger;

public class Validation {
	final Logger logger = Logger.getLogger (Validation.class);
	
	public boolean checkNullString(String strInput) {
		if(strInput == null || strInput.length() == 0) {
			return false;
		}
		return true;
	}
	
	public boolean checkNullUploadFile(HttpServletRequest request) throws IOException, ServletException{
		
		for (Part part : request.getParts()) {
			//System.out.println("name = " + part.getName());
			//System.out.println("size = " + part.getSize());
			//System.out.println("filename = " + part.getSubmittedFileName());
			if(part.getSubmittedFileName() == "" || part.getSize() == 0){
				return false;	//uploadファイルなし
			}
			break;
		}		
		return true;	//uploadファイルあり
	}
}
