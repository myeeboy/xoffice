package com.hg.xs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.util.UUID;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.jacob.com.ComThread;

/**
 * OfficeServlet
 * 
 * @author wanghg
 */
public class XOfficeServlet extends HttpServlet {
	private static final long serialVersionUID = 280698472747919447L;
	private static Logger logger = Logger.getLogger(XOfficeServlet.class.getName());
	private String key = null;
	private boolean clearTempFile = true;
	public void init() throws ServletException {
		super.init();
		this.key = nvl(this.getInitParameter("key"), "");
		if ("false".equals(this.getInitParameter("clearTempFile"))) {
			this.clearTempFile = false;
		}
		ComThread.InitMTA();
		logger.info("Office Servlet Init!");
	}
		
	private String nvl(String str, String def) {
		return str != null ? str : def;
	}
	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws IOException {
		doAct(request, response, false);
	}
	public void doPost(HttpServletRequest request, HttpServletResponse response)
			throws IOException {
		doAct(request, response, true);
	}
	public void doAct(HttpServletRequest request, HttpServletResponse response, boolean post)
			throws IOException {
		String reqkey = nvl(request.getParameter("_key"), "");
		if (!reqkey.equals(this.key)) {
			logger.info("Invalid '_key'!");
			return;
		}
		String format = request.getHeader("_xformat");
		if (format == null) {
			format = request.getParameter("_xformat");
		}
		if (format == null) {
			logger.info("No Paramater '_xformat'!");
			return;
		}
		String urlFile = null;
		if (!post) {
			urlFile = request.getParameter("_file");
			if (urlFile == null) {
				logger.info("No Paramater '_file'!");
				return;
			}
		}
		long s = System.currentTimeMillis();
		logger.info("Convert start...");			
		String fileName = System.getProperty("java.io.tmpdir");
		if (!fileName.endsWith("/") && !fileName.endsWith("\\")) {
			fileName += "/";
		}
		fileName += UUID.randomUUID().toString();
		File src = new File(fileName + "." + format);
		File tar = new File(fileName + ".pdf");
		logger.info(src.getAbsolutePath() + " >>> " + tar.getAbsolutePath());
		InputStream in = null;
		OutputStream out = null;
		try {
			if (post) {
				in = request.getInputStream();
				out = new FileOutputStream(src);
				pipe(in, out);
			} else {
				in = new URL(urlFile).openStream();
				out = new FileOutputStream(src);
				pipe(in, out);
			}
			if (format.startsWith("doc")) {
				WordApp.toPdf(src.getAbsolutePath(), tar.getAbsolutePath());
			} else if (format.startsWith("xls")) {
				ExcelApp.toPdf(src.getAbsolutePath(), tar.getAbsolutePath());
			} else if (format.startsWith("ppt")) {
				PowerPointApp.toPdf(src.getAbsolutePath(), tar.getAbsolutePath());
			}
			out = response.getOutputStream();
			in = new FileInputStream(tar);
			pipe(in, out);
		} catch (Throwable t) {
			response.setStatus(500);
			logger.log(Level.SEVERE, t.getMessage(), t);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (Exception e) {
				}
			}
			if (out != null) {
				try {
					out.close();
				} catch (Exception e) {
				}
			}
			if (clearTempFile) {
				src.delete();
				tar.delete();
			}
		}
		logger.info("Convert stop,Use " + (System.currentTimeMillis() - s) + " ms!");
	}

	private static void pipe(InputStream in, OutputStream out)
			throws IOException {
		int len;
		byte[] buf = new byte[4096];
		while (true) {
			len = in.read(buf);
			if (len > 0) {
				out.write(buf, 0, len);
			} else {
				break;
			}
		}
		in.close();
		out.flush();
		out.close();
	}
}
