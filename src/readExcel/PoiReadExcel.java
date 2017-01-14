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
		String filePath = "file/2016-3考勤数据.xlsx";
		calculateResult(filePath);
		// readXml(filePath);
		// Calendar cal = Calendar.getInstance();
		// cal.set(year, month, 0);
		// int year = cal.get(Calendar.YEAR);
		// int m = cal.get(Calendar.MONTH);
		// int dayNumOfMonth = Export.getDaysByYearMonth(year, m + 1);
		// cal.set(Calendar.DAY_OF_MONTH, 1);// 从一号开始
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
		// if (!status.contains("早上")) {
		// writeZS(df, key);
		// }
		// if (!status.contains("中午")) {
		// writeZW(df, key);
		// }
		// if (!status.contains("下午")) {
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
		cal.set(Calendar.DAY_OF_MONTH, 1);// 从一号开始
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
					if (!status.contains("早上")) {
						writeZS(df, key);
					}
					if (!status.contains("中午")) {
						writeZW(df, key);
					}
					if (!status.contains("下午")) {
						writeXW(df, key);
					}

				}
			}
		}
		// }
	}

	public static void writeZS(String day, String name) {
		File zs = new File("D:/早上没打卡.txt");
		try {
			BufferedWriter be = new BufferedWriter(new FileWriter(zs, true));
			System.out.println(name + "\t" + day + "\t早上沒打卡");
			be.write(name + "\t" + day + "\t早上沒打卡");
			be.newLine();
			be.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void writeZW(String day, String name) {
		File zw = new File("D:/中午没打卡.txt");
		try {
			BufferedWriter be = new BufferedWriter(new FileWriter(zw, true));
			System.out.println(name + "\t" + day + "\t中午沒打卡");
			be.write(name + "\t" + day + "\t中午沒打卡");
			be.newLine();
			be.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void writeXW(String day, String name) {
		File xw = new File("D:/下午没打卡.txt");
		try {
			BufferedWriter be = new BufferedWriter(new FileWriter(xw, true));
			System.out.println(name + "\t" + day + "\t下午沒打卡");
			be.write(name + "\t" + day + "\t下午沒打卡");
			be.newLine();
			be.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void writeQueQin(String day, String name) {
		File qq = new File("D:/缺勤.txt");
		try {
			BufferedWriter be = new BufferedWriter(new FileWriter(qq, true));
			System.out.println(name + "\t" + day + "\t缺勤");
			be.write(name + "\t" + day + "\t缺勤");
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
			sta = "早上正常";
		else if (DBUtil.compare(time, m2) > 0 && DBUtil.compare(time, n1) < 0 && DBUtil.closer(m2, time, n1))
			sta = "早上迟到";
		else if (DBUtil.compare(time, m2) > 0 && DBUtil.compare(time, n1) < 0 && !DBUtil.closer(m2, time, n1))
			sta = "中午早退";
		else if (DBUtil.compare(time, n1) >= 0 && DBUtil.compare(time, n2) <= 0)
			sta = "中午正常";
		else if (DBUtil.compare(time, n2) > 0 && DBUtil.compare(time, a) < 0 && DBUtil.closer(n2, time, a))
			sta = "中午迟到";
		else if (DBUtil.compare(time, n2) > 0 && DBUtil.compare(time, a) < 0 && !DBUtil.closer(n2, time, a))
			sta = "下午早退";
		else if (DBUtil.compare(time, a) >= 0)
			sta = "下午正常";
		return sta;
	}

	public static void readXml(String fileName) {
		boolean isE2007 = false; // 判断是否是excel2007格式
		if (fileName.endsWith("xlsx"))
			isE2007 = true;
		try {
			InputStream input = new FileInputStream(fileName); // 建立输入流
			Workbook wb = null;
			// 根据文件格式(2003或者2007)来初始化
			if (isE2007)
				wb = new XSSFWorkbook(input);
			else
				wb = new HSSFWorkbook(input);
			Sheet sheet = wb.getSheetAt(0); // 获得第一个表单
			Iterator<Row> rows = sheet.rowIterator(); // 获得第一个表单的迭代器
			while (rows.hasNext()) {
				Row row = rows.next(); // 获得行数据
				// System.out.println("Row #" + row.getRowNum()); // 获得行号从0开始
				if (row.getRowNum() > 0) {
					Iterator<Cell> cells = row.cellIterator(); // 获得第一行的迭代器
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
