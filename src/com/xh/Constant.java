package com.xh;

public class Constant {
	// 可以操作的文件类型
	public static final String FILE_SUFFIX[] = { ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt" };
	// 转换得出的文件类型
	public static final String PDF_SUFFIX = ".pdf";
	// OpenOffice安装路径
	public static final String OFFICE_HOME = "D:/OpenOffice/OpenOffice.v4";
	// 设置转换端口，默认为8100
	public static final int PORT = 8100;
	// 设置任务执行超时为2分钟
	public static final long EXECUTE_OVERTIME = 2 * 60 * 1000;
	// 设置任务队列超时为5分钟
	public static final long QUEUE_OVERTIME = 5 * 60 * 1000;


}
