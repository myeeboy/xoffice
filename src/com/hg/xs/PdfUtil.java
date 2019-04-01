package com.hg.xs;

import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.lowagie.text.DocumentException;
import com.lowagie.text.Element;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.pdf.PdfContentByte;
import com.lowagie.text.pdf.PdfGState;
import com.lowagie.text.pdf.PdfReader;
import com.lowagie.text.pdf.PdfStamper;

/**
 * PdfUtil
 * @author wanghg
 */
public class PdfUtil {
	private static Logger logger = Logger.getLogger(PdfUtil.class.getName());

	/**
	 * 添加水印
	 * @param in
	 * @param out
	 * @param watermark
	 * @throws IOException 
	 */
	public static void addWatermark(InputStream in, OutputStream out,
			String watermark) throws IOException {
		try {
			PdfReader reader = new PdfReader(in);    
			PdfStamper stamper = new PdfStamper(reader, out);
	        BaseFont font = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H",BaseFont.NOT_EMBEDDED, true);
	        Rectangle pageSize = null;  
	        PdfGState gs = new PdfGState();  
	        gs.setFillOpacity(0.1f);  
	        gs.setStrokeOpacity(0.4f);  
	        int pageCount = reader.getNumberOfPages() + 1;
	        Graphics2D g = (Graphics2D) new BufferedImage(10, 10, BufferedImage.TYPE_INT_RGB).getGraphics();
	        int textWidth = (int) new Font("宋体", Font.PLAIN, 24).getStringBounds(watermark, g.getFontRenderContext()).getWidth();;
	        PdfContentByte pcb;
	        for (int i = 1; i < pageCount; i++) {   
	        	pageSize = reader.getPageSizeWithRotation(i);   
	        	int wn = (int) Math.ceil(pageSize.getWidth() / textWidth);
	        	int hn = (int) Math.ceil(pageSize.getHeight() / textWidth);
	            pcb = stamper.getOverContent(i);   
	            pcb.saveState();  
	            pcb.setGState(gs);  
	            pcb.beginText();    
	            pcb.setFontAndSize(font, 24);    
	            for (int m = 0; m < hn; m++) {    
	                for (int n = 0; n < wn; n++) {  
	                	pcb.showTextAligned(Element.ALIGN_LEFT , watermark, 24 + n * textWidth, m * textWidth, 45);
	                }  
	            }  
	            pcb.endText();    
	        }
	        stamper.close();  
	        reader.close();
		} catch (DocumentException e) {
			logger.log(Level.SEVERE, e.getMessage(), e);
		}    
	}
}
