package hoge;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConvertServlet extends HttpServlet {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		System.out.println("START");
		
		
		FileInputStream in = null;
	    Workbook wb = null;
	    
	    
	    //既存のworkbookを開く
	    try{
	      in = new FileInputStream("/Users/aa352872/Desktop/sample.xlsx");
	      wb = WorkbookFactory.create(in);
	    }catch(IOException e){
	      System.out.println(e.toString());
	    }catch(InvalidFormatException e){
	      System.out.println(e.toString());
	    }finally{
	      try{
	        in.close();
	      }catch (IOException e){
	        System.out.println(e.toString());
	      }
	    }

	    Sheet sheet = wb.createSheet("new sheet");
	    
	    //workbookを別名で保存
	    FileOutputStream out = null;
	    try{
	      out = new FileOutputStream("/Users/aa352872/Desktop/sample4_1.xlsx");
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
