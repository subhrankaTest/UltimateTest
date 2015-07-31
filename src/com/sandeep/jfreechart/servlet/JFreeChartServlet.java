package com.sandeep.jfreechart.servlet;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.image.RenderedImage;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

import javax.imageio.ImageIO;
import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartRenderingInfo;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.entity.StandardEntityCollection;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;

/**
 * Servlet implementation class JFreeChartServlet
 */
public class JFreeChartServlet extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public JFreeChartServlet() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		
		
		//Reading and writing into the excel
		WriteExcelDemo writeobj=new WriteExcelDemo();
		ReadExcelDemo readObj=new ReadExcelDemo();	
		
		WriteExcelDemo.writeExcel(request, response);
		ReadExcelDemo.readExcel(request, response);
		
		//reading and writing into the excel
		
		
		
		
		
		
		response.setContentType("image/png");
		
        ServletOutputStream os = response.getOutputStream();
        OutputStream out = response.getOutputStream();

        DefaultPieDataset dataset = new DefaultPieDataset();
		dataset.setValue("High", 23);
		dataset.setValue("Low", 32);
		dataset.setValue("Critical", 24);
		dataset.setValue("medium", 40);

		boolean legend = true;
		boolean tooltips = false;
		boolean urls = false;

		JFreeChart chart = ChartFactory.createPieChart("Defects", dataset, legend, tooltips, urls);
		

		chart.setBorderPaint(Color.GREEN);
		chart.setBorderStroke(new BasicStroke(5.0f));
		chart.setBorderVisible(true);
		
		response.setContentType("image/png");

        ChartUtilities.writeChartAsPNG(out, chart, 300, 300);
		
		try {

            final ChartRenderingInfo info = new ChartRenderingInfo(new StandardEntityCollection());
            final File file1 = new File(request.getContextPath() + "/images/piechart/pie.png");

            ChartUtilities.saveChartAsPNG(file1, chart, 600, 400, info);
        } catch (Exception e) {
            System.out.println(e);

        } 
        
        
        
       // RenderedImage  chartImage = chart.createBufferedImage(300, 300);
       // ImageIO.write(chartImage, "png", os);
        
        os.flush();
        os.close();
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		
			}

}
