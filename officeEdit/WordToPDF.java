package com.repast.core.util;

import java.io.File;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class WordToPDFUtil {

	public static void main(String[] args) {
		WordToPDFUtil d = new WordToPDFUtil();
		d.wordToPDF("D:/bbb.docx", "D:/ddd.pdf");
	}

	static final int wdFormatPDF = 17;

	public void wordToPDF(String sfileName, String toFileName) {
		long start = System.currentTimeMillis();
		ComThread.InitSTA();// 必须先启动winword.exe进程
		ActiveXComponent app = null;
		Dispatch doc = null;

		try {
			app = new ActiveXComponent("Word.Application");
			app.setProperty("Visible", new Variant(false));
			Dispatch docs = app.getProperty("Documents").toDispatch();
			doc = Dispatch.call(docs, "Open", sfileName).toDispatch();
			System.out.println("打开文档..." + sfileName);
			System.out.println("转换文档到PDF..." + toFileName);
			File tofile = new File(toFileName);
			if (tofile.exists()){
				tofile.delete(); 
			}
			Dispatch.call(doc, "SaveAs", toFileName,wdFormatPDF); 
			long end = System.currentTimeMillis(); 
			System.out.println("转换完成..用时：" + (end - start) + "ms.");
		} catch (Exception e) {
			System.out.println("========Error:文档转换失败：" + e.getMessage()); 
		} finally {
			Dispatch.call(doc, "Close", false); 
			System.out.println("关闭文档"); 
			if (app != null) 
				app.invoke("Quit", new Variant[] {}); 
		}

		ComThread.Release();
	}
}
