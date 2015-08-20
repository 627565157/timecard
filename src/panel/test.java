package panel;

public class test {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		String path = "C:\\Users\\Sin\\Desktop\\10月份考勤数据.xls";
		String newPath = path.replace('\\', '/');
		System.out.println(newPath);
	}

}
