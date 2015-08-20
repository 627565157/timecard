package panel;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import db.DBUtil;

public class Export {
	public static String date;

	public Export() {

	}

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		long start = System.currentTimeMillis();
		Export ex = new Export();
		try {
			ex.export();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		System.out.println("finish...");
		long end = System.currentTimeMillis();
		System.out.println("����ʱ�䣺" + (end - start) + "����");
	}

	public static int getDaysByYearMonth(int year, int month) {

		Calendar a = Calendar.getInstance();
		a.set(Calendar.YEAR, year);
		a.set(Calendar.MONTH, month - 1);
		a.set(Calendar.DATE, 1);
		a.roll(Calendar.DATE, -1);
		int maxDate = a.get(Calendar.DATE);
		return maxDate;
	}

	public static List<String> getDate() {
		List<String> list = new ArrayList<String>();
		// Calendar ca = Calendar.getInstance();
		// int year = ca.get(Calendar.YEAR);
		// int month = ca.get(Calendar.MONTH);
		// int day = getDaysByYearMonth(year, month);
		// for (int i = 1; i <= day; i++) {
		// StringBuilder sb = new StringBuilder();
		// sb.append(Integer.toString(year));
		// if (month < 10)
		// sb.append("-0" + Integer.toString(month));
		// else
		// sb.append("-" + Integer.toString(month));
		// if (i < 10)
		// sb.append("-0" + Integer.toString(i));
		// else
		// sb.append("-" + Integer.toString(i));
		// list.add(sb.toString());
		// }
		String[] date2 = date.split("-");
		String year = date2[0];
		String month = date2[1];
		boolean b = checkYear(year);
		list = getDay(year, month, b);
//		System.out.println("��������begin�� ");
//		for(String s : list) {
//			System.out.println(s);
//		}
//		System.out.println("��������end;");
		return list;
	}

	public static List<String> getDay(String year, String month, boolean b) {
		List<String> list = new ArrayList<String>();
		switch (Integer.parseInt(month)) {
		case 1:
			for (int i = 1; i < 32; i++) {
				StringBuilder sb = new StringBuilder();
				if (i < 10)
					sb.append(year + "-01-0" + i);
				else
					sb.append(year + "-01-" + i);
				list.add(sb.toString());
			}
			break;
		case 2:
			if (b) {
				for (int i = 1; i < 30; i++) {
					StringBuilder sb = new StringBuilder();
					if (i < 10)
						sb.append(year + "-02-0" + i);
					else
						sb.append(year + "-02-" + i);
					list.add(sb.toString());
				}
			} else {
				for (int i = 1; i < 29; i++) {
					StringBuilder sb = new StringBuilder();
					if (i < 10)
						sb.append(year + "-02-0" + i);
					else
						sb.append(year + "-02-" + i);
					list.add(sb.toString());
				}
			}
			break;
		case 3:
			for (int i = 1; i < 32; i++) {
				StringBuilder sb = new StringBuilder();
				if (i < 10)
					sb.append(year + "-03-0" + i);
				else
					sb.append(year + "-03-" + i);
				list.add(sb.toString());
			}
			break;
		case 4:
			for (int i = 1; i < 31; i++) {
				StringBuilder sb = new StringBuilder();
				if (i < 10)
					sb.append(year + "-04-0" + i);
				else
					sb.append(year + "-04-" + i);
				list.add(sb.toString());
			}
			break;
		case 5:
			for (int i = 1; i < 32; i++) {
				StringBuilder sb = new StringBuilder();
				if (i < 10)
					sb.append(year + "-05-0" + i);
				else
					sb.append(year + "-05-" + i);
				list.add(sb.toString());
			}
			break;
		case 6:
			for (int i = 1; i < 31; i++) {
				StringBuilder sb = new StringBuilder();
				if (i < 10)
					sb.append(year + "-06-0" + i);
				else
					sb.append(year + "-06-" + i);
				list.add(sb.toString());
			}
			break;
		case 7:
			for (int i = 1; i < 32; i++) {
				StringBuilder sb = new StringBuilder();
				if (i < 10)
					sb.append(year + "-07-0" + i);
				else
					sb.append(year + "-07-" + i);
				list.add(sb.toString());
			}
			break;
		case 8:
			for (int i = 1; i < 32; i++) {
				StringBuilder sb = new StringBuilder();
				if (i < 10)
					sb.append(year + "-08-0" + i);
				else
					sb.append(year + "-08-" + i);
				list.add(sb.toString());
			}
			break;
		case 9:
			for (int i = 1; i < 31; i++) {
				StringBuilder sb = new StringBuilder();
				if (i < 10)
					sb.append(year + "-09-0" + i);
				else
					sb.append(year + "-09-" + i);
				list.add(sb.toString());
			}
			break;
		case 10:
			for (int i = 1; i < 32; i++) {
				StringBuilder sb = new StringBuilder();
				if (i < 10)
					sb.append(year + "-10-0" + i);
				else
					sb.append(year + "-10-" + i);
				list.add(sb.toString());
			}
			break;
		case 11:
			for (int i = 1; i < 31; i++) {
				StringBuilder sb = new StringBuilder();
				if (i < 10)
					sb.append(year + "-11-0" + i);
				else
					sb.append(year + "-11-" + i);
				list.add(sb.toString());
			}
			break;
		case 12:
			for (int i = 1; i < 32; i++) {
				StringBuilder sb = new StringBuilder();
				if (i < 10)
					sb.append(year + "-12-0" + i);
				else
					sb.append(year + "-12-" + i);
				list.add(sb.toString());
			}
			break;
		default:
			break;
		}
		return list;
	}

	public static boolean checkYear(String s) {
		int n = Integer.parseInt(s);
		return (n % 4 == 0 && n % 100 != 0 || n % 400 == 0);
	}

	public static List<String> getAllID() {
		List<String> list = new ArrayList<String>();
		Connection conn = DBUtil.getConnection();
		Statement state = null;
		ResultSet rs;
		try {
			state = conn.createStatement();
			rs = state.executeQuery("select ���ں��� from kq");
			while (rs.next()) {
				String s = rs.getString(1);
				if (!list.contains(s)) {
					list.add(s);
				}
			}
			rs.close();
			state.close();
			conn.close();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return list;
	}

	private String getNameById(String userId) throws SQLException {
		Connection conn = DBUtil.getConnection();
		Statement state = null;
		ResultSet rs = null;
		String name = "";
		try {
			state = conn.createStatement();
			rs = state.executeQuery("select ���� from kq where ���ں��� ='" + userId
					+ "'");
			if (rs.next()) {
				name = rs.getString(1);
			}
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			rs.close();
			state.close();
			conn.close();
		}
		return name;

	}

	public boolean export() throws SQLException {
		boolean b = false;
		List<String> ID = getAllID();
		List<String> date = getDate();
		File qq = new File("D:/ȱ��.txt");
		File zs = new File("D:/����û��.txt");
		File zw = new File("D:/����û��.txt");
		File xw = new File("D:/����û��.txt");
		Connection conn = DBUtil.getConnection();
		Statement state = conn.createStatement();
		String name = "";
		for (String id : ID) {
			name = getNameById(id);
			for (String day : date) {
				List<String> list = new ArrayList<String>();
				List<String> list2 = new ArrayList<String>();
				ResultSet rs = null;
				try {
					rs = state.executeQuery("select ״̬ from kq where ���ں���='"
							+ id + "' and ��������='" + day + "'");
					while (rs.next()) {
						list.add(rs.getString(1));
					}
					rs.close();
				} catch (SQLException e) {
					e.printStackTrace();
				}
				if (list.size() == 0)
					list2.add("ȱ��");
				else {
//					for (String s : list) {
//						System.out.println("test :" + s);
//					}
//					System.out
//							.println("����û�� ����� "
//									+ (!list.contains("��������") && !list
//											.contains("��������")));
					if (!list.contains("��������") && !list.contains("���ϳٵ�")) {
						list2.add("����û��");
					}
					if (!list.contains("��������") && !list.contains("��������")
							&& !list.contains("����ٵ�")) {
						list2.add("����û��");
					}
					if (!list.contains("��������") && !list.contains("��������")) {
						list2.add("����û��");
					}

				}
				if (list2.contains("ȱ��")) {
					try {
						BufferedWriter be = new BufferedWriter(new FileWriter(
								qq, true));
						be.write(id + "\t" + name + "\t" + day + "\tȱ��");
						be.newLine();
						be.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
				} else {
					try {
						BufferedWriter be = new BufferedWriter(
								new OutputStreamWriter(new FileOutputStream(zs,
										true)));
						BufferedWriter be2 = new BufferedWriter(
								new OutputStreamWriter(new FileOutputStream(zw,
										true)));
						BufferedWriter be3 = new BufferedWriter(
								new OutputStreamWriter(new FileOutputStream(xw,
										true)));
						if (list2.contains("����û��")) {
							be.flush();
							be.write(id + "\t" + name + "\t" + day
									+ "\t����û��\n");
							be.newLine();
							System.out.println(id + "\t" + name + "\t" + day
									+ "\t����û��\n");
						}
						if (list2.contains("����û��")) {
							be2.flush();
							be2.write(id + "\t" + name + "\t" + day
									+ "\t����û��\n");
							be2.newLine();
							System.out.println(id + "\t" + name + "\t" + day
									+ "\t����û��\n");
						}
						if (list2.contains("����û��")) {
							be3.flush();
							be3.write(id + "\t" + name + "\t" + day
									+ "\t����û��\n");
							System.out.println(id + "\t" + name + "\t" + day
									+ "\t����û��\n");
							be3.newLine();
						}
						be.close();
						be2.close();
						be3.close();
					} catch (Exception e) {
						e.printStackTrace();
					}
				}
			}
		}
		state.close();
		conn.close();
		b = true;
		return b;
	}

}
