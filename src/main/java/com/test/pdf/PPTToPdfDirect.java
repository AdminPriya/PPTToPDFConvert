package com.test.pdf;

import java.awt.Dimension;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;


import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.PdfGraphics2D;
import com.lowagie.text.pdf.PdfWriter;

public class PPTToPdfDirect {

	public static void main(String[] args) throws IOException, DocumentException {

		//load any ppt file
		FileInputStream inputStream = new FileInputStream("F:\\file_example_PPT_250kB.ppt");
		SlideShow ppt = new SlideShow(inputStream);
		inputStream.close();
		Dimension pgsize = ppt.getPageSize();


		//take first slide and draw it directly into PDF via awt.Graphics2D interface.
		//Slide slide = ppt.getSlides()[0];

		Slide slide = ppt.getSlides()[0];

		Document document = new Document();

     
			PdfWriter pdfWriter = PdfWriter.getInstance(document, new FileOutputStream("F:\\PPTtoPDF12.pdf"));
			document.setPageSize(new Rectangle((float)pgsize.getWidth(), (float) pgsize.getHeight()));
			document.open();

			PdfGraphics2D graphics = (PdfGraphics2D) pdfWriter.getDirectContent().createGraphics((float)pgsize.getWidth(), (float)pgsize.getHeight());
			slide.draw(graphics);

			graphics.dispose();
	
		document.close();

	}

}
