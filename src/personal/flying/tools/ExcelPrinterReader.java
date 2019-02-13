package personal.flying.tools;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ExcelPrinterReader {

	public static void main(String[] args) throws Exception {
		try {
			
			System.out.println("========start=========");
			System.out.println("");
			File outputTemplateFile = new File("C:/excelreader/ģ��.xls");
			ExcelOperator o = new ExcelOperator();
			Map<String, String> noMapping = o.readNoMapping("C:/excelreader/���.xlsx");
		
			getFileList("C:/excelreader/source");
			for (File sourceFile : fileList) {
				OutputExcelData outputData = o.readSourceExcel(sourceFile.getAbsolutePath());
				if(outputData == null)
					continue;
				
				String personNo = noMapping.get(outputData.identityNo);
				if (personNo == null || personNo.trim().length() == 0) {
					System.out.println("δ�ҵ����˵ı��:" + outputData.name);					
				} else {
					outputData.personNo = personNo;
				}
				String destFilePath = getDestFileName(sourceFile.getAbsolutePath());
				File destFile = new File(destFilePath);
				destFile.getParentFile().mkdirs();
				destFile.delete();
				Files.copy(outputTemplateFile.toPath(), destFile.toPath());
				o.writeToDestExcel(destFilePath, outputData);
				System.out.println(outputData.name + " Finish");
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println("");
		System.out.println("");
		System.out.println("==========end=============");
		System.out.println("Press any key to exit");
		System.in.read();
	}

	public static String getDestFileName(String sourceFile) {
		return "C:/excelreader/output/"
				+ sourceFile.substring("C:/excelreader/source/".length(), sourceFile.length() - 4) + "_output.xls";
	}

	private static List<File> fileList = new ArrayList<File>(2000);

	public static void getFileList(String strPath) {
		File dir = new File(strPath);
		File[] files = dir.listFiles(); // ���ļ�Ŀ¼���ļ�ȫ����������
		if (files != null) {
			for (int i = 0; i < files.length; i++) {
				String fileName = files[i].getName();
				if (files[i].isDirectory()) { // �ж����ļ������ļ���
					getFileList(files[i].getAbsolutePath()); // ��ȡ�ļ�����·��
				} else if (fileName.endsWith(".xls")) { // �ж��ļ����Ƿ���.avi��β
					String strFileName = files[i].getAbsolutePath();
					fileList.add(files[i]);
				} else {
					continue;
				}
			}

		}
	}
}
