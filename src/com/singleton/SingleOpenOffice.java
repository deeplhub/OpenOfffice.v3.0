package com.singleton;

import java.io.File;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

import com.xh.Constant;

/**
 * 
 * @author H.Yang
 * @date 2017��9��8��
 */
public class SingleOpenOffice {

	private static SingleOpenOffice start = new SingleOpenOffice();
	private static OfficeManager officeManager;

	// ��ȡΨһ���õĶ���
	public static SingleOpenOffice getStart() {
		DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
		System.out.println("׼����������....");
		configuration.setOfficeHome(FileSuffixType.FILEPATH.getName()); // ����OpenOffice.org��װĿ¼
		configuration.setPortNumber((int) FileSuffixType.PORT.getValue()); // ����ת���˿ڣ�Ĭ��Ϊ8100
		configuration.setTaskExecutionTimeout(Long.valueOf(String.valueOf(FileSuffixType.EXECUTE_OVERTIME.getValue())));
		configuration.setTaskQueueTimeout(Long.valueOf(String.valueOf(FileSuffixType.QUEUE_OVERTIME.getValue())));

		officeManager = configuration.buildOfficeManager();
		officeManager.start(); // ��������
		System.out.println("office���������ɹ�!");
		return start;
	}

	/**
	 * �ĵ�ת��
	 * <hr>
	 * ��doc,docx,xls,xlsx,ppt,pptx,txt���ĵ�ת����PDF�ĵ��������ָ�������ַĬ�ϵ�ǰ�ļ���ַ
	 * 
	 * @author H.Yang
	 * @date 2016��12��13��
	 * @explain
	 * 
	 * @param inputFilePath
	 *            - ת���ļ���ַ�����룩
	 * @param outputFilePath
	 *            - �����ַ���ɿգ�
	 * @param newFileName
	 *            - ���ļ������ɿգ�
	 * @return
	 */
	public String execute2Pdf(String inputFilePath, String outputFilePath, String newFileName) {
		File inputFile = new File(inputFilePath);
		String fileName = inputFile.getName();
		String prefix = fileName.substring(fileName.lastIndexOf(".") + 0);
		String outputPath = null;

		boolean isTrue = false;
		if (!inputFile.exists()) {
			System.out.println("�ļ�������!");
			return null;
		}

		for (String name : Constant.FILE_SUFFIX) {
			if (fileName.endsWith(name)) {
				isTrue = true;
				break;
			}
		}

		if (!isTrue) {
			System.out.println("�ļ���ʽ����");
			return null;
		}

		if (outputFilePath != null) {
			outputPath = newFileName == null ? outputFilePath + fileName.replace(prefix, Constant.PDF_SUFFIX) : outputFilePath
					+ fileName.replace(fileName, newFileName) + Constant.PDF_SUFFIX;
		} else {
			outputPath = newFileName == null ? inputFile.getPath().replace(prefix, Constant.PDF_SUFFIX) : inputFile.getPath().replace(fileName, newFileName)
					+ Constant.PDF_SUFFIX;
		}

		File outputFile = new File(outputPath);
		if (!outputFile.exists()) {
			// ִ�з���������
			execute(inputFile, outputFile);
		} else {
			System.out.println("�ļ��Ѵ���");
		}
		return outputPath;
	}

	/**
	 * ִ�з���������
	 * 
	 * @author H.Yang
	 * @date 2016��12��13��
	 * @explain
	 * 
	 * @param inputFile
	 * @param outputFile
	 */
	private static void execute(File inputFile, File outputFile) {
		long startTime = System.currentTimeMillis();// ��ȡ��ʼʱ��

		try {
			System.out.println("�����ĵ�ת��ת��:" + inputFile + " --> " + outputFile);
			OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
			converter.convert(inputFile, outputFile);
			System.out.println("Officeת���ɹ�");
		} catch (Exception e) {
			getStop();
			e.printStackTrace();
		}

		long endTime = System.currentTimeMillis(); // ��ȡ����ʱ��
		System.out.println("��������ʱ�䣺 " + (endTime - startTime) / 1000 + "s");
	}

	public static void getStop() {
		if (officeManager != null) {
			officeManager.stop();
		}
		System.out.println("office�رճɹ�!");
	}

}
