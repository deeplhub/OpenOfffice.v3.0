package com.singleton;

/**
 * 
 * @author M.Dong
 * @date 2017年9月8日
 */
public enum FileSuffixType {

	DOC(".doc", "Word文档"), //
	DOCX(".docx", "Word文档"), //
	XLS(".xls", "Excel文档"), //
	XLSX(".xlsx", "Excel文档"), //
	PPT(".ppt", "PowerPoint文档"), //
	PPTX(".pptx", "PowerPoint文档"), //
	TXT(".text", "记事本文档"), //
	RETURNTYPE(".pdf", "便携式文档格式"), //
	FILEPATH("D:/OpenOffice/OpenOffice.v4", "OpenOffice安装路径"), //
	PORT(8100, "设置转换端口，默认为8100"), //
	EXECUTE_OVERTIME(2 * 60 * 1000, "设置任务执行超时为2分钟"), //
	QUEUE_OVERTIME(5 * 60 * 1000, "设置任务队列超时为5分钟");

	private Object value;
	private String name;
	private String explain;

	FileSuffixType(String name, String explain) {
		this.name = name;
		this.explain = explain;
	}

	FileSuffixType(Object value, String explain) {
		this.value = value;
		this.explain = explain;
	}

	public Object getValue() {
		return value;
	}

	public String getName() {
		return name;
	}

	public String getExplain() {
		return explain;
	}

}
