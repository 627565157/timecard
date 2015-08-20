package readExcel;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class Read {

	public static void main(String[] args) {
		List<List<String>> lists = getAll("C:/Users/Sin/Desktop/10月份考勤数据.xls");
		for(List<String> list : lists) {
			for(String s : list) {
				System.out.print(s+"\t");
			}
			System.out.println();
		}
	}

	public static List<List<String>> getAll(String url) 
	{
		jxl.Workbook readwb = null;
		List<List<String>> list = new ArrayList<List<String>>();

		try

		{

			// 构建Workbook对象, 只读Workbook对象

			// 直接从本地文件创建Workbook

			InputStream instream = new FileInputStream(url);

			readwb = Workbook.getWorkbook(instream);

			// Sheet的下标是从0开始

			// 获取第一张Sheet表

			Sheet readsheet = readwb.getSheet(0);

			// 获取Sheet表中所包含的总列数

			int rsColumns = readsheet.getColumns();

			// 获取Sheet表中所包含的总行数

			int rsRows = readsheet.getRows();

			// 获取指定单元格的对象引用

			for (int i = 1; i < rsRows; i++)

			{
				List<String> row = new ArrayList<String>();
				for (int j = 0; j < rsColumns; j++) {
					Cell cell = readsheet.getCell(j, i);
					row.add(cell.getContents());
				}
				list.add(row);
				// System.out.println();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;

	}
}
