package com.hg.xs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;

/**
 * XOffice客户端
 * @author coffish
 * https://view.xdocin.com
 */
public class XOfficeClient {
	/**
	 * 服务地址
	 */
	private String url;
	/**
	 * 服务口令
	 */
	private String key;
	/**
	 * 水印
	 */
	private String watermark;
	/**
	 * 服务地址
	 * @return
	 */
	public String getUrl() {
		return url;
	}
	/**
	 * 服务地址
	 * @param url
	 */
	public void setUrl(String url) {
		this.url = url;
	}
	/**
	 * 账号口令
	 * @return
	 */
	public String getKey() {
		return key;
	}
	/**
	 * 获取水印
	 * @return
	 */
	public String getWatermark() {
		return watermark;
	}
	/**
	 * 设置水印
	 * @param watermark
	 */
	public void setWatermark(String watermark) {
		this.watermark = watermark;
	}
	/**
	 * 构造器
	 * @param url 服务地址
	 */
	public XOfficeClient(String url) {
		this(url, "");
	}
	/**
	 * 构造器
	 * @param url 服务地址
	 * @param key 账号
	 */
	public XOfficeClient(String url, String key) {
		this.url = url;
		this.key = key;
	}
	/**
	 * office文件转pdf文件
	 * @param officeFile doc/docx/xls/xlsx/ppt/pptx
	 * @param pdfFile pdf文件，如：a.pdf
	 * @throws IOException
	 */
	public void to(File officeFile, File pdfFile) throws IOException {
		String format = "";
		String name = officeFile.getName();
		if (name.indexOf('.') > 0) {
			format = name.substring(name.lastIndexOf('.') + 1).toLowerCase();
		}
		FileInputStream fin = new FileInputStream(officeFile);
		FileOutputStream fout = new FileOutputStream(pdfFile);
		try {			
			to(fin, format, fout);
		} finally {
			fin.close();
			fout.close();
		}
	}
	/**
	 * office文件URL（http）转pdf文件
	 * @param officeFileUrl
	 * @param format 文件格式：doc/docx/xls/xlsx/ppt/pptx
	 * @param pdfFile pdf文件，如：a.pdf
	 * @throws IOException
	 */
	public void to(String officeFileUrl, String format, File pdfFile) throws IOException {
		FileOutputStream fout = new FileOutputStream(pdfFile);
		try {			
			to(officeFileUrl, format, fout);
		} finally {
			fout.close();
		}
	}
	/**
	 * office文件URL（http）转pdf文件
	 * @param officeFileUrl
	 * @param format 文件格式：doc/docx/xls/xlsx/ppt/pptx
	 * @param pdfOut pdf流
	 * @throws IOException
	 */
	public void to(String officeFileUrl, String format, OutputStream pdfOut) throws IOException {
		StringBuffer sb = new StringBuffer();
		sb.append(this.url).append("/xoffice");
		sb.append("?_key=").append(encode(this.key));
		sb.append("&_file=").append(encode(officeFileUrl));
		sb.append("&_xformat=").append(format);
		if (this.watermark != null) {
			sb.append("&_watermark=").append(encode(this.watermark));
		}
		URL url = new URL(sb.toString());
		pipe(url.openStream(), pdfOut);
	}
	/**
	 * office文件流转pdf文件流
	 * @param officeFileIn office文件流
	 * @param format 文件格式：doc/docx/xls/xlsx/ppt/pptx
	 * @param pdfOut pdf文件流
	 * @throws IOException
	 */
	public void to(InputStream officeFileIn, String format, OutputStream pdfOut) throws IOException {
		StringBuffer sb = new StringBuffer();
		sb.append(this.url).append("/xoffice");
		sb.append("?_key=").append(encode(this.key));
		sb.append("&_xformat=").append(format);
		if (this.watermark != null) {
			sb.append("&_watermark=").append(encode(this.watermark));
		}
		HttpURLConnection httpConn = (HttpURLConnection) new URL(sb.toString()).openConnection();
		httpConn.setRequestProperty("Content-Type", "application/octet-stream");
		httpConn.setDoOutput(true);
		OutputStream reqOut = httpConn.getOutputStream();
		pipe(officeFileIn, reqOut);
		pipe(httpConn.getInputStream(), pdfOut);
	}
	public static void main(String[] args) {
//		XOfficeClient xc = new XOfficeClient("http://localhost:9090");
//		try {
//			xc.to(new File("c:/tmp.docx"), new File("c:/tmp.pdf"));
//		} catch (IOException e) {
//			e.printStackTrace();
//		}
	}
	private static String encode(String str) {
        try {
            return URLEncoder.encode(str, "UTF-8");
        } catch (UnsupportedEncodingException e) {
            return str;
        }
    }
    private static void pipe(InputStream in, OutputStream out) throws IOException {
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
        out.flush();
        out.close();
        in.close();
    }
}