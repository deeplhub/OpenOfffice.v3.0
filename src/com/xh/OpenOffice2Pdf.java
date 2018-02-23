package com.xh;

import java.io.File;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

public class OpenOffice2Pdf {

	private static OfficeManager officeManager;

	/**
	 * 文档转换
	 * <hr>
	 * 将doc,docx,xls,xlsx,ppt,pptx,txt等文档转换成PDF文档，如果不指定输出地址默认当前文件地址
	 * 
	 * @author H.Yang
	 * @date 2016年12月13日
	 * @explain
	 * 
	 * @param inputFilePath
	 *            - 转换文件地址（必须）
	 * @param outputFilePath
	 *            - 输出地址（可空）
	 * @param newFileName
	 *            - 新文件名（可空）
	 * @return
	 */
	public String execute2Pdf(String inputFilePath, String outputFilePath, String newFileName) {

		return processing(inputFilePath, outputFilePath, newFileName);
	}

	/**
	 * 数据加工
	 * 
	 * @author H.Yang
	 * @date 2016年12月13日
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
			System.out.println("文件不存在!");
			return null;
		}

		for (String name : Constant.FILE_SUFFIX) {
			if (fileName.endsWith(name)) {
				isTrue = true;
				break;
			}
		}
		
		if (!isTrue) {
			System.out.println("文件格式错误");
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
			//执行方法服务功能
			execute(inputFile, outputFile);
		} else {
			System.out.println("文件已存在");
		}

		return outputPath;
	}

	/**
	 * 执行方法服务功能
	 * 
	 * @author H.Yang
	 * @date 2016年12月13日
	 * @explain
	 * 
	 * @param inputFile
	 * @param outputFile
	 */
	private static void execute(File inputFile, File outputFile) {
		long startTime = System.currentTimeMillis();// 获取开始时间

		try {
			if (startService()) {
				System.out.println("进行文档转换转换:" + inputFile + " --> " + outputFile);
				OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
				converter.convert(inputFile, outputFile);
				System.out.println("Office转换成功");
				stopService();
			}
		} catch (Exception e) {
			stopService();
			e.printStackTrace();
		}

		long endTime = System.currentTimeMillis(); // 获取结束时间
		System.out.println("程序运行时间： " + (endTime - startTime) / 1000 + "s");
	}

	/**
	 * 开启服务
	 * 
	 * @author H.Yang
	 * @date 2016年12月13日
	 * @explain
	 * 
	 * @return
	 */
	private static boolean startService() {
		DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
		try {
			// 设置任务执行超时,
			long executeOvertime = Constant.EXECUTE_OVERTIME;
			// 设置任务队列超时
			long queueOvertime = Constant.QUEUE_OVERTIME;

			System.out.println("准备启动服务....");
			configuration.setOfficeHome(Constant.OFFICE_HOME); // 设置OpenOffice.org安装目录
			configuration.setPortNumber(Constant.PORT); // 设置转换端口，默认为8100
			configuration.setTaskExecutionTimeout(executeOvertime);
			configuration.setTaskQueueTimeout(queueOvertime);

			officeManager = configuration.buildOfficeManager();
			officeManager.start(); // 启动服务
			System.out.println("office服务启动成功!");
			return true;
		} catch (Exception ce) {
			System.out.println("office转换服务启动失败!详细信息:" + ce);
			return false;
		}
	}

	/**
	 * 关闭服务
	 * 
	 * @author H.Yang
	 * @date 2016年12月13日
	 * @explain
	 *
	 */
	private static void stopService() {
		System.out.println("准备关闭服务....");
		if (officeManager != null) {
			officeManager.stop();
		}
		System.out.println("office关闭成功!");
	}

}
