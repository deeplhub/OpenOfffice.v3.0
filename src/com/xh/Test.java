package com.xh;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 
 * @author Hyang
 * @date 2016Äê12ÔÂ14ÈÕ
 * @explain
 */
public class Test {

	public static void main(String[] args) {
		
		OpenOffice2Pdf openOffice2Pdf = new OpenOffice2Pdf();
		String path = "E:/xxxxxx.doc";

		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmm");
		String dateTime = sdf.format(new Date());
		String resultPath = openOffice2Pdf.execute2Pdf(path, null, dateTime);
		System.out.println(resultPath);
	}
}
