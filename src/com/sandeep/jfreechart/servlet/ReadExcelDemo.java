package com.sandeep.jfreechart.servlet;

import java.awt.BasicStroke;
import java.awt.Color;
import java.util.HashMap;
import java.util.Iterator;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.data.general.DefaultPieDataset;

import java.io.*;

//import statements
public class ReadExcelDemo extends HttpServlet 
{	
	
	

    /**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public static void readExcel(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		 
    {
        try
        {
            FileInputStream file = new FileInputStream(new File(request.getContextPath()) +"src\\xls\\howtodoinjava_demo.xls");
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
 
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            HashMap<String, String> sevierityVal = new HashMap<String, String>();
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) 
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                 
                while (cellIterator.hasNext()) 
                {
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    switch (cell.getCellType()) 
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print(cell.getNumericCellValue());
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue());
                            break;
                    }
                }
                System.out.println("");
            }
            
            
            DefaultPieDataset dataset = new DefaultPieDataset( );
            dataset.setValue("High", new Double( 20 ) );
            dataset.setValue("Medium", new Double( 20 ) );
            dataset.setValue("Low", new Double( 40 ) );
            dataset.setValue("Critical", new Double( 10 ) );

            JFreeChart chart = ChartFactory.createPieChart(
               "Defect status", // chart title
               dataset, // data
               true, // include legend
               true,
               false);
               
            int width = 640; /* Width of the image */
            int height = 480; /* Height of the image */ 
           // File pieChart = new File( "C:/Users/Administrator/workspace/TestExcel/src/PieChart.jpeg" );
            final File file1 = new File(request.getContextPath() + "/images/PieChartRead.jpeg");

           // ChartUtilities.saveChartAsJPEG( pieChart , chart , width , height );


            
            file.close();
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
    }
}
}
