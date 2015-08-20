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
		List<List<String>> lists = getAll("C:/Users/Sin/Desktop/10�·ݿ�������.xls");
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

			// ����Workbook����, ֻ��Workbook����

			// ֱ�Ӵӱ����ļ�����Workbook

			InputStream instream = new FileInputStream(url);

			readwb = Workbook.getWorkbook(instream);

			// Sheet���±��Ǵ�0��ʼ

			// ��ȡ��һ��Sheet��

			Sheet readsheet = readwb.getSheet(0);

			// ��ȡSheet������������������

			int rsColumns = readsheet.getColumns();

			// ��ȡSheet������������������

			int rsRows = readsheet.getRows();

			// ��ȡָ����Ԫ��Ķ�������

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
