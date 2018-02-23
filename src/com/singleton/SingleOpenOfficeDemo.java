package com.singleton;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 
 * @author H.Yang
 * @date 2017Äê9ÔÂ8ÈÕ
 */
public class SingleOpenOfficeDemo {
	public static void main(String[] args) {
		SingleOpenOffice openOffice = SingleOpenOffice.getStart();

		String path = "E:/XXXXXXXXX.doc";

		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmm");
		String dateTime = sdf.format(new Date());
		String resultPath = openOffice.execute2Pdf(path, null, dateTime);
		System.out.println(resultPath);

		openOffice.getStop();
	}
}
