package hoge;

import java.io.IOException;
import java.util.List;

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
		boolean f = false;
		
		for (Part part : request.getParts()) {
			if(part != null){
				f = true;
			}
		}		
		return f;
	}
}
