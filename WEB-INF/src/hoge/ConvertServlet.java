package hoge;

import java.io.FileOutputStream;
import java.io.IOException;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConvertServlet extends HttpServlet {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		
		Workbook wb = new XSSFWorkbook();
//		XSSFWorkbook wb = new XSSFWorkbook();
//		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("new sheet");
		
		//create the new workbook
		System.out.println("START");
		FileOutputStream out = null;
	    try{
	      out = new FileOutputStream("/Users/aa352872/Desktop/sampleDayoooooooooo2_1.xlsx");
	      wb.write(out);
	    }catch(IOException e){
	      System.out.println(e.toString());
	    }finally{
	      try {
	        out.close();
	      }catch(IOException e){
	        System.out.println(e.toString());
	      }
	    }
		
	    System.out.println("END\n-------");
		RequestDispatcher dispatcher = request.getRequestDispatcher("/jsp/Result.jsp");
		dispatcher.forward(request, response);
	}
	
	
	
}
