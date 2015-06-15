import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.Iterator;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class hssf extends HttpServlet {

	/**
	 * Constructor of the object.
	 */
	public hssf() {
		super();
	}

	/**
	 * Destruction of the servlet. <br>
	 */
	public void destroy() {
		super.destroy(); // Just puts "destroy" string in log
		// Put your code here
	}

	/**
	 * The doGet method of the servlet. <br>
	 *
	 * This method is called when a form has its tag value method equals to get.
	 * 
	 * @param request the request send by the client to the server
	 * @param response the response send by the server to the client
	 * @throws ServletException if an error occurred
	 * @throws IOException if an error occurred
	 */
	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		
		String f=request.getParameter("text1");
		log(f);

		response.setContentType("text/html");
		PrintWriter out = response.getWriter();
		out.println("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\">");
		out.println("<HTML>");
		out.println("  <HEAD><TITLE>A Servlet</TITLE></HEAD>");
		out.println("  <BODY>");
		out.print("    This is ");
		out.print(this.getClass());
		out.println(", using the GET method");
		out.println("  </BODY>");
		out.println("</HTML>");
		out.flush();
		out.close();
	}

	/**
	 * The doPost method of the servlet. <br>
	 *
	 * This method is called when a form has its tag value method equals to post.
	 * 
	 * @param request the request send by the client to the server
	 * @param response the response send by the server to the client
	 * @throws ServletException if an error occurred
	 * @throws IOException if an error occurred
	 */
	public void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {

		response.setContentType("text/html");
		PrintWriter out = response.getWriter();
		out.println("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\">");
		out.println("<HTML>");
		out.println("  <HEAD><TITLE>A Servlet</TITLE></HEAD>");
		out.println("  <BODY>");
		out.print("    This is ");
		out.print(this.getClass());
		out.println(", using the POST method");
		out.println("  </BODY>");
		out.println("</HTML>");
		out.flush();
		out.close();
	}

	/**
	 * Initialization of the servlet. <br>
	 *
	 * @throws ServletException if an error occurs
	 */
	public void init() throws ServletException {
		
	
		String path="d://1.xls";
		// Put your code here
		//try {
			//excel();

		//} catch (IOException e) {
			// TODO Auto-generated catch block
		//	e.printStackTrace();
		//}
		
		try {
			readTable();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void excel() throws IOException {
		// TODO Auto-generated method stub
		
		Workbook wb = new HSSFWorkbook(); 
        //������һ��sheet��ҳ��������Ϊ new sheet 
        Sheet sheet = wb.createSheet("new sheet"); 
        Row row = sheet.createRow((short) 0); 
        // ��row���ϴ���һ������ 
        Cell cell = row.createCell(0); 
        //���÷������ʾ 
        cell.setCellValue(1); 
 
        // Or do it on one line. 
        row.createCell(1).setCellValue(1.2); 
        row.createCell(2).setCellValue("This is a string �ٶȷ�������"); 
        row.createCell(3).setCellValue(true); 
 
        //����һ���ļ� ����Ϊworkbook.xls 
        
        String path="e://workbook.xls";
        FileOutputStream fileOut = new FileOutputStream(path); 
        // �����洴���Ĺ�����������ļ��� 
        log("excel"+path);
        wb.write(fileOut); 
        //�ر������ 
        fileOut.close(); 
	}
	
    //ͨ���Ե�Ԫ���������ʽ����ȡ��Ϣ ������Ҫ�жϵ�Ԫ������Ͳſ���ȡ��ֵ  
    public static void readTable() throws Exception{   
        InputStream ips=new FileInputStream("e://workbook.xls");   
        HSSFWorkbook wb=new HSSFWorkbook(ips);   
        HSSFSheet sheet=wb.getSheetAt(0);   
        for(Iterator ite=sheet.rowIterator();ite.hasNext();){   
            HSSFRow row=(HSSFRow)ite.next();   
            System.out.println();   
            for(Iterator itet=row.cellIterator();itet.hasNext();){   
                HSSFCell cell=(HSSFCell)itet.next();   
                switch(cell.getCellType()){     
                case HSSFCell.CELL_TYPE_BOOLEAN:     
                    //�õ�Boolean����ķ���     
                    System.out.print(cell.getBooleanCellValue()+" ");     
                    break;     
                case HSSFCell.CELL_TYPE_NUMERIC:     
                    //�ȿ��Ƿ������ڸ�ʽ     
                    if(HSSFDateUtil.isCellDateFormatted(cell)){     
                        //��ȡ���ڸ�ʽ     
                        System.out.print(cell.getDateCellValue()+" ");     
                    }else{     
                        //��ȡ����     
                        System.out.print(cell.getNumericCellValue()+" ");     
                    }     
                    break;     
                case HSSFCell.CELL_TYPE_FORMULA:     
                    //��ȡ��ʽ     
                    System.out.print(cell.getCellFormula()+" ");     
                    break;     
                case HSSFCell.CELL_TYPE_STRING:     
                    //��ȡString     
                    System.out.print(cell.getRichStringCellValue().toString()+" ");     
                    break;                       
            }     
            }   
        }   
           
    }   
       
    //ֱ�ӳ�ȡexcel�е�����   
    public static void extractor() throws Exception{   
        InputStream ips=new FileInputStream("d://test2-1.xls");   
        HSSFWorkbook wb=new HSSFWorkbook(ips);   
        ExcelExtractor ex=new ExcelExtractor(wb);   
        ex.setFormulasNotResults(true);   
        ex.setIncludeSheetNames(false);   
        String text=ex.getText();   
        System.out.println(text);    
    }  
	
	

}
