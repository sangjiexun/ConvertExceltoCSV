package hoge;

import org.apache.log4j.Logger;

public class Validation {
	final Logger logger = Logger.getLogger (Validation.class);
	
	public boolean checkNullString(String strInput) {
		if(strInput == null || strInput.length() == 0) {
			return false;
		}
		return true;
	}	
}
