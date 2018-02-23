package com.xh;

import java.io.File;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

public class OpenOffice2Pdf {

	private static OfficeManager officeManager;

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

		return processing(inputFilePath, outputFilePath, newFileName);
	}

	/**
	 * ���ݼӹ�
	 * 
	 * @author H.Yang
	 * @date 2016��12��13��
	 * @explain
	 * 
	 * @return
	 */
	private static String processing(String inputFilePath, String outputFilePath, String newFileName) {
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
			outputPath = newFileName == null ? outputFilePath + fileName.replace(prefix, Constant.PDF_SUFFIX)
					: outputFilePath + fileName.replace(fileName, newFileName) + Constant.PDF_SUFFIX;
		} else {
			outputPath = newFileName == null ? inputFile.getPath().replace(prefix, Constant.PDF_SUFFIX)
					: inputFile.getPath().replace(fileName, newFileName) + Constant.PDF_SUFFIX;
		}
		
		File outputFile = new File(outputPath);
		if (!outputFile.exists()) {
			//ִ�з���������
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
			if (startService()) {
				System.out.println("�����ĵ�ת��ת��:" + inputFile + " --> " + outputFile);
				OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
				converter.convert(inputFile, outputFile);
				System.out.println("Officeת���ɹ�");
				stopService();
			}
		} catch (Exception e) {
			stopService();
			e.printStackTrace();
		}

		long endTime = System.currentTimeMillis(); // ��ȡ����ʱ��
		System.out.println("��������ʱ�䣺 " + (endTime - startTime) / 1000 + "s");
	}

	/**
	 * ��������
	 * 
	 * @author H.Yang
	 * @date 2016��12��13��
	 * @explain
	 * 
	 * @return
	 */
	private static boolean startService() {
		DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
		try {
			// ��������ִ�г�ʱ,
			long executeOvertime = Constant.EXECUTE_OVERTIME;
			// ����������г�ʱ
			long queueOvertime = Constant.QUEUE_OVERTIME;

			System.out.println("׼����������....");
			configuration.setOfficeHome(Constant.OFFICE_HOME); // ����OpenOffice.org��װĿ¼
			configuration.setPortNumber(Constant.PORT); // ����ת���˿ڣ�Ĭ��Ϊ8100
			configuration.setTaskExecutionTimeout(executeOvertime);
			configuration.setTaskQueueTimeout(queueOvertime);

			officeManager = configuration.buildOfficeManager();
			officeManager.start(); // ��������
			System.out.println("office���������ɹ�!");
			return true;
		} catch (Exception ce) {
			System.out.println("officeת����������ʧ��!��ϸ��Ϣ:" + ce);
			return false;
		}
	}

	/**
	 * �رշ���
	 * 
	 * @author H.Yang
	 * @date 2016��12��13��
	 * @explain
	 *
	 */
	private static void stopService() {
		System.out.println("׼���رշ���....");
		if (officeManager != null) {
			officeManager.stop();
		}
		System.out.println("office�رճɹ�!");
	}

}
