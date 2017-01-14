package readExcel;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import db.DBUtil;
import panel.Export;

public class PoiReadExcel {
	public static List<String> listKey = new ArrayList<String>();
	public static Map<String, String> map = new LinkedHashMap<String, String>();
	public static int year = -1;
	public static int month = -1;

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		String filePath = "file/2016-3��������.xlsx";
		calculateResult(filePath);
		// readXml(filePath);
		// Calendar cal = Calendar.getInstance();
		// cal.set(year, month, 0);
		// int year = cal.get(Calendar.YEAR);
		// int m = cal.get(Calendar.MONTH);
		// int dayNumOfMonth = Export.getDaysByYearMonth(year, m + 1);
		// cal.set(Calendar.DAY_OF_MONTH, 1);// ��һ�ſ�ʼ
		// for (int i = 0; i < dayNumOfMonth; i++, cal.add(Calendar.DATE, 1)) {
		// Date d = cal.getTime();
		// SimpleDateFormat simpleDateFormat = new
		// SimpleDateFormat("yyyy-MM-dd");
		// String df = simpleDateFormat.format(d);
		// for (String key : listKey) {
		// String mapKey = key + "_" + df;
		// if (map.get(mapKey) == null) {
		// writeQueQin(df, key);
		// } else {
		// String[] kaoTimes = map.get(mapKey).split("_");
		// System.out.println(mapKey + " " + map.get(mapKey));
		// String status = "";
		// for (String time : kaoTimes) {
		// status += checkStatus(time);
		// }
		// if (!status.contains("����")) {
		// writeZS(df, key);
		// }
		// if (!status.contains("����")) {
		// writeZW(df, key);
		// }
		// if (!status.contains("����")) {
		// writeXW(df, key);
		// }
		//
		// }
		// }
		// }
	}

	public static void calculateResult(String filePath) {
		readXml(filePath);
		Calendar cal = Calendar.getInstance();
		cal.set(year, month, 0);
		int year = cal.get(Calendar.YEAR);
		int m = cal.get(Calendar.MONTH);
		int dayNumOfMonth = Export.getDaysByYearMonth(year, m + 1);
		cal.set(Calendar.DAY_OF_MONTH, 1);// ��һ�ſ�ʼ
		List<String> days = new ArrayList<String>();
		for (int i = 0; i < dayNumOfMonth; i++, cal.add(Calendar.DATE, 1)) {
			Date d = cal.getTime();
			SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
			String df = simpleDateFormat.format(d);
			days.add(df);
		}
		// for (int i = 0; i < dayNumOfMonth; i++, cal.add(Calendar.DATE, 1)) {
		// Date d = cal.getTime();
		// SimpleDateFormat simpleDateFormat = new
		// SimpleDateFormat("yyyy-MM-dd");
		// String df = simpleDateFormat.format(d);
		for (String key : listKey) {
			for (String df : days) {
				String mapKey = key + "_" + df;
				if (map.get(mapKey) == null) {
					writeQueQin(df, key);
				} else {
					String[] kaoTimes = map.get(mapKey).split("_");
					System.out.println(mapKey + "   " + map.get(mapKey));
					String status = "";
					for (String time : kaoTimes) {
						status += checkStatus(time);
					}
					if (!status.contains("����")) {
						writeZS(df, key);
					}
					if (!status.contains("����")) {
						writeZW(df, key);
					}
					if (!status.contains("����")) {
						writeXW(df, key);
					}

				}
			}
		}
		// }
	}

	public static void writeZS(String day, String name) {
		File zs = new File("D:/����û��.txt");
		try {
			BufferedWriter be = new BufferedWriter(new FileWriter(zs, true));
			System.out.println(name + "\t" + day + "\t���ϛ]��");
			be.write(name + "\t" + day + "\t���ϛ]��");
			be.newLine();
			be.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void writeZW(String day, String name) {
		File zw = new File("D:/����û��.txt");
		try {
			BufferedWriter be = new BufferedWriter(new FileWriter(zw, true));
			System.out.println(name + "\t" + day + "\t����]��");
			be.write(name + "\t" + day + "\t����]��");
			be.newLine();
			be.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void writeXW(String day, String name) {
		File xw = new File("D:/����û��.txt");
		try {
			BufferedWriter be = new BufferedWriter(new FileWriter(xw, true));
			System.out.println(name + "\t" + day + "\t����]��");
			be.write(name + "\t" + day + "\t����]��");
			be.newLine();
			be.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void writeQueQin(String day, String name) {
		File qq = new File("D:/ȱ��.txt");
		try {
			BufferedWriter be = new BufferedWriter(new FileWriter(qq, true));
			System.out.println(name + "\t" + day + "\tȱ��");
			be.write(name + "\t" + day + "\tȱ��");
			be.newLine();
			be.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static String checkStatus(String time) {
//		String m1 = "6:00:00";
//		String m2 = "8:00:00";
//		String n1 = "11:55:00";
//		String n2 = "13:30:00";
//		String a = "17:30:00";
		String m1 = "4:00:00";
		String m2 = "8:00:00";
		String n1 = "12:00:00";
		String n2 = "13:00:00";
		String a = "17:00:00";
		String sta = "";
		if (DBUtil.compare(time, m1) > 0 && DBUtil.compare(time, m2) <= 0)
			sta = "��������";
		else if (DBUtil.compare(time, m2) > 0 && DBUtil.compare(time, n1) < 0 && DBUtil.closer(m2, time, n1))
			sta = "���ϳٵ�";
		else if (DBUtil.compare(time, m2) > 0 && DBUtil.compare(time, n1) < 0 && !DBUtil.closer(m2, time, n1))
			sta = "��������";
		else if (DBUtil.compare(time, n1) >= 0 && DBUtil.compare(time, n2) <= 0)
			sta = "��������";
		else if (DBUtil.compare(time, n2) > 0 && DBUtil.compare(time, a) < 0 && DBUtil.closer(n2, time, a))
			sta = "����ٵ�";
		else if (DBUtil.compare(time, n2) > 0 && DBUtil.compare(time, a) < 0 && !DBUtil.closer(n2, time, a))
			sta = "��������";
		else if (DBUtil.compare(time, a) >= 0)
			sta = "��������";
		return sta;
	}

	public static void readXml(String fileName) {
		boolean isE2007 = false; // �ж��Ƿ���excel2007��ʽ
		if (fileName.endsWith("xlsx"))
			isE2007 = true;
		try {
			InputStream input = new FileInputStream(fileName); // ����������
			Workbook wb = null;
			// �����ļ���ʽ(2003����2007)����ʼ��
			if (isE2007)
				wb = new XSSFWorkbook(input);
			else
				wb = new HSSFWorkbook(input);
			Sheet sheet = wb.getSheetAt(0); // ��õ�һ����
			Iterator<Row> rows = sheet.rowIterator(); // ��õ�һ�����ĵ�����
			while (rows.hasNext()) {
				Row row = rows.next(); // ���������
				// System.out.println("Row #" + row.getRowNum()); // ����кŴ�0��ʼ
				if (row.getRowNum() > 0) {
					Iterator<Cell> cells = row.cellIterator(); // ��õ�һ�еĵ�����
					Cell personIdCell = row.getCell(0);
					String personId = personIdCell.getStringCellValue();
					Cell personNameCell = row.getCell(1);
					String personName = personNameCell.getStringCellValue();
					Cell dayCell = row.getCell(2);
					Date day = HSSFDateUtil.getJavaDate(dayCell.getNumericCellValue());
					Cell dayTimeCell = row.getCell(3);
					Date dayTime = HSSFDateUtil.getJavaDate(dayTimeCell.getNumericCellValue());
					if (!listKey.contains(personName)) {
						listKey.add(personName);
					}
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
					if (year == -1) {
						String[] yearMonthDay = sdf.format(day).toString().split("-");
						year = Integer.parseInt(yearMonthDay[0]);
						month = Integer.parseInt(yearMonthDay[1]);
					}
					String key = personName + "_" + sdf.format(day).toString();
					SimpleDateFormat sdf2 = new SimpleDateFormat("HH:mm:ss");
					if (map.containsKey(key)) {
						String dates = (String) map.get(key);
						String newDates = dates + "_" + sdf2.format(dayTime).toString();
						map.put(key, newDates);
					} else {
						map.put(key, sdf2.format(dayTime).toString());
					}
				}
			}
		} catch (IOException ex) {
			ex.printStackTrace();
		}
	}

}
