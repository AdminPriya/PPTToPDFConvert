package com.test.pdf;


import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;


public class PPTtoPNG {

	// Converting the slides of a PPT into Images using Java
	

		public static void main(String args[])
			throws IOException
		{

			// create an empty presentation
			FileInputStream inputStream = new FileInputStream("F:\\file_example_PPT_250kB.ppt");
			SlideShow ppt = new SlideShow(inputStream);


			// get the dimension and size of the slide
			Dimension pgsize = ppt.getPageSize();
			Slide[] slide = ppt.getSlides();
			BufferedImage img = null;

			System.out.println(slide.length);

			for (int i = 0; i < slide.length; i++) {
				img = new BufferedImage(
					pgsize.width, pgsize.height,
					BufferedImage.TYPE_INT_RGB);
				Graphics2D graphics = img.createGraphics();

				// clear area
				graphics.setPaint(Color.white);
				graphics.fill(new Rectangle2D.Float(
					0, 0, pgsize.width, pgsize.height));

				// draw the images
				slide[i].draw(graphics);
				FileOutputStream out = new FileOutputStream("F:\\img.png");
				javax.imageio.ImageIO.write(img, "png", out);
				ppt.write(out);
				out.close();
				System.out.println(i);
			}
		}
	}



