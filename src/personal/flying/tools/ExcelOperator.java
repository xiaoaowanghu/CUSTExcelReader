package personal.flying.tools;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperator {
	private static final String EXCEL_XLS = "xls";
	private static final String EXCEL_XLSX = "xlsx";

	public ExcelOperator() {
	}

	public Map<String, String> readNoMapping(String mappingPath) {
		Map<String, String> result = new HashMap<String, String>();
		Workbook workbook = null;
		int count = 0;
		try {
			File excelFile = new File(mappingPath); // �����ļ�����
			checkExcelVaild(excelFile);
			workbook = getWorkbok(excelFile);
			Sheet sheet = workbook.getSheetAt(0); // ������һ��Sheet

			for (Row row : sheet) {
				try {
					count++;
					// Ϊ������һ��Ŀ¼����count
					if (count == 1) {
						continue;
					}
					// �����ǰ��û�����ݣ�����ѭ��
					if (row.getCell(0).toString().equals("")) {
						return result;
					}
					String identityNo = getValue(row.getCell(4)).toString().trim();
					String personNo = getValue(row.getCell(5)).toString().trim();
					result.put(identityNo, personNo);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		} catch (Exception error) {
			System.out.println("readNoMapping error. filename=" + mappingPath);
			error.printStackTrace();
			return null;
		} finally {
			try {
				if (workbook != null)
					workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		return result;
	}

	public OutputExcelData readSourceExcel(String filePath) {
		Workbook workbook = null;
		OutputExcelData result = new OutputExcelData();
		result.infos = new ArrayList<>(6);
		FileInputStream in = null;
		int count = 0;
		try {
			File excelFile = new File(filePath); // �����ļ�����
//			in = new FileInputStream(excelFile); // �ļ���
			checkExcelVaild(excelFile);
			workbook = getWorkbok(excelFile);
			Sheet sheet = workbook.getSheetAt(0); // ������һ��Sheet

			for (Row row : sheet) {
				count++;
				// Ϊ������һ��Ŀ¼����count
				if (count == 1) {
					continue;
				}

				TaxInfo info = new TaxInfo();
				info.name = getValue(row.getCell(1)).toString().trim();
				info.identityType = getValue(row.getCell(2)).toString().trim();
				info.identityNo = getValue(row.getCell(3)).toString().trim();
				info.taxStartDate = getValue(row.getCell(6)).toString().trim();
				info.taxEndDate = getValue(row.getCell(7)).toString().trim();
				info.inputAmount = getValue(row.getCell(8)).toString().trim();
				info.taxPaidAmount = getValue(row.getCell(30)).toString().trim();
				result.identityNo = info.identityNo;

				if (result.name == null)
					result.name = info.name;
				else if (result.name != null && !info.name.equals(result.name))
					throw new Exception("ͬһ�ű����ֲ���ͬ");

				result.infos.add(info);
			}
			if (result.infos.size() > 6) {
				throw new Exception("��¼����6��");
			}
		} catch (Exception error) {
			System.out.println("readSourceExcel error. fileName=" + filePath);
			error.printStackTrace();
			return null;
		} finally {
			try {
				if (workbook != null)
					workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		return result;
	}

	public void writeToDestExcel(String filePath, OutputExcelData excelData) {
		Workbook workBook = null;
		OutputStream out = null;
		try {
			File excelFile = new File(filePath);
			workBook = getWorkbok(excelFile);
			Sheet sheet = workBook.getSheetAt(0);

			Row noRow = sheet.getRow(1);
			noRow.getCell(6).setCellValue(excelData.personNo);

			int dataCount = excelData.infos.size();
			for (int j = 0; j < dataCount; j++) {
				TaxInfo taxInfo = excelData.infos.get(j);
				// ����һ�У��ӵڶ��п�ʼ������������
				Row row = sheet.getRow(3 + j);
				// �õ�Ҫ�����ÿһ����¼
				row.getCell(0).setCellValue(taxInfo.name);
				row.getCell(1).setCellValue(taxInfo.identityType);
				row.getCell(2).setCellValue(taxInfo.identityNo);
				row.getCell(3).setCellValue(taxInfo.taxStartDate);
				row.getCell(4).setCellValue(taxInfo.taxEndDate);
				row.getCell(5).setCellValue(taxInfo.inputAmount);
				row.getCell(6).setCellValue(taxInfo.taxPaidAmount);
			}
			// �����ļ��������������ӱ����������У���������sheet�������κβ�����������Ч
			out = new FileOutputStream(filePath);
			workBook.write(out);
			out.flush();
		} catch (Exception e) {
			System.out.println("writeToDestExcel error. fileName=" + filePath + " itemCount=" + excelData.infos.size());
			e.printStackTrace();
		} finally {
			try {
				if (out != null)
					out.close();
				if (workBook != null)
					workBook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	/**
	 * �ж�Excel�İ汾,��ȡWorkbook
	 */
	private Workbook getWorkbok(File file) throws IOException {
		Workbook wb = null;
		InputStream in = new FileInputStream(file); // �ļ���
		if (file.getName().toLowerCase().endsWith(EXCEL_XLS)) { // Excel 2003
			wb = new HSSFWorkbook(in);
		} else if (file.getName().toLowerCase().endsWith(EXCEL_XLSX)) { // Excel 2007/2010
			wb = new XSSFWorkbook(in);
		}
		return wb;
	}

	/**
	 * �ж��ļ��Ƿ���excel
	 */
	private void checkExcelVaild(File file) throws Exception {
		if (!file.exists()) {
			throw new Exception("�ļ�������");
		}
		if (!(file.isFile() && (file.getName().endsWith(EXCEL_XLS) || file.getName().endsWith(EXCEL_XLSX)))) {
			throw new Exception("�ļ�����Excel");
		}
	}

	private Object getValue(Cell cell) {
		if (cell == null)
			return "";

		Object obj = null;
		switch (cell.getCellTypeEnum()) {
		case BOOLEAN:
			obj = cell.getBooleanCellValue();
			break;
		case ERROR:
			obj = cell.getErrorCellValue();
			break;
		case NUMERIC:
			obj = cell.getNumericCellValue();
			break;
		case STRING:
			obj = cell.getStringCellValue();
			break;
		default:
			break;
		}

		if (obj == null) {
			// logger.debug("object get from cell(" + cell.getRowIndex() + ", " +
			// cell.getColumnIndex() + ") is null");
			obj = "";
		}

		return obj;
	}
}
