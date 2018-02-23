package com.singleton;

/**
 * 
 * @author M.Dong
 * @date 2017��9��8��
 */
public enum FileSuffixType {

	DOC(".doc", "Word�ĵ�"), //
	DOCX(".docx", "Word�ĵ�"), //
	XLS(".xls", "Excel�ĵ�"), //
	XLSX(".xlsx", "Excel�ĵ�"), //
	PPT(".ppt", "PowerPoint�ĵ�"), //
	PPTX(".pptx", "PowerPoint�ĵ�"), //
	TXT(".text", "���±��ĵ�"), //
	RETURNTYPE(".pdf", "��Яʽ�ĵ���ʽ"), //
	FILEPATH("D:/OpenOffice/OpenOffice.v4", "OpenOffice��װ·��"), //
	PORT(8100, "����ת���˿ڣ�Ĭ��Ϊ8100"), //
	EXECUTE_OVERTIME(2 * 60 * 1000, "��������ִ�г�ʱΪ2����"), //
	QUEUE_OVERTIME(5 * 60 * 1000, "����������г�ʱΪ5����");

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
