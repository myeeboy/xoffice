package com.hg.xs;

import java.util.logging.Level;
import java.util.logging.Logger;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * WordApp
 * @author wanghg
 */
public class PowerPointApp {
	private static Logger logger = Logger.getLogger(WordApp.class.getName());

	public static void toPdf(String src, String tar) {
		ActiveXComponent app = null;
		Dispatch doc = null;
		try {
			app = new ActiveXComponent("PowerPoint.Application");
			app.setProperty("AutomationSecurity", new Variant(3));
			Dispatch docs = app.getProperty("Presentations").toDispatch();
			doc = Dispatch.call(
					docs,
					"Open",
					new Object[] { src, Boolean.TRUE, Boolean.TRUE,
							Boolean.FALSE }).toDispatch();
			Dispatch.call(doc, "SaveAs",
					new Object[] { tar, Integer.valueOf(32) });
		} finally {
			if (doc != null) {
				try {
					Dispatch.call(doc, "Close");
				} catch (Throwable t) {
					logger.log(Level.SEVERE, t.getMessage(), t);
				}
			}
			if (app != null) {
				try {
					app.invoke("Quit");
					app.safeRelease();
				} catch (Throwable t) {
					logger.log(Level.SEVERE, t.getMessage(), t);
				}
			}
		}
	}
}
