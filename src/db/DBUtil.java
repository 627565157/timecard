package db;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.List;


public class DBUtil {
	public DBUtil(){
		
	}

	public static void main(String[] args) throws SQLException {
		// TODO Auto-generated method stub
//		List list=Read.getAll();
//		DBUtil db=new DBUtil();
//		db.update();
		getConnection();
        System.out.println("finish...");
	}
	
	
	public static Connection getConnection()
	{
		Connection conn = null;
		try {
		    Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
		    String url = "jdbc:odbc:Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=./db/cs.accdb";
		    conn = DriverManager.getConnection(url);
		} catch (SQLException e) {
		    System.out.println("SQL Exception: "+ e.toString());
		} catch (ClassNotFoundException cE) {
		    System.out.println("Class Not Found Exception: "+ cE.toString());
		}
		return conn;
	}
	
	public int insert(List<List<String>> list)
	{
		Connection conn=getConnection();
		PreparedStatement pstate;
		int count=0;
		try {
			pstate = conn.prepareStatement("insert into kq(���ں���,����,��������,����ʱ��) values(?,?,?,?)");
			for(List<String> list2:list)
			{
				pstate.setString(1, list2.get(0).toString());
				pstate.setString(2, list2.get(1).toString());
				pstate.setString(3, list2.get(2).toString());
				pstate.setString(4, list2.get(3).toString());
				pstate.addBatch();
				count++;
			}
			pstate.executeBatch();
			pstate.close();
			conn.close();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return count;
	}
	
	//�Ƚ������ַ���ʱ��Ĵ�С
	public static int compare(String s1,String s2)
	{
		int i1=Integer.parseInt(s1.replace(":", "").toString());
		int i2=Integer.parseInt(s2.replace(":", "").toString());
		return i1-i2;
	}
	
	//�Ƚ�s2����s1��s3�ĸ�����
	public static boolean closer(String s1,String s2,String s3)
	{
		boolean b=false;
		double d1=Integer.parseInt(s1.split(":")[0])+Double.parseDouble(s1.split(":")[1])/60;
		double d2=Integer.parseInt(s2.split(":")[0])+Double.parseDouble(s2.split(":")[1])/60;
		double d3=Integer.parseInt(s3.split(":")[0])+Double.parseDouble(s3.split(":")[1])/60;
		if((d2-d1)<=(d3-d2))
		   b=true;
		return b;
	}

	public int update()
	{
		String m1="6:00:00";
		String m2="8:00:00";
		String n1="11:55:00";
		String n2="13:30:00";
		String a="17:30:00";
		int count=0;
		Connection conn = null;
		try {
		    Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
		    String url = "jdbc:odbc:Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=./db/cs.accdb";
		    conn = DriverManager.getConnection(url);
		    Statement state=conn.createStatement();
		    ResultSet rs=state.executeQuery("select ID,����ʱ�� from kq");
		    while(rs.next())
		    {
		    	String sta="";
		    	int id=rs.getInt(1);
		    	String time=rs.getString(2);
		    	if(compare(time,m1)>0&&compare(time,m2)<=0)
		    		sta="��������";
		    	else if(compare(time,m2)>0&&compare(time,n1)<0&&closer(m2,time,n1))
		    		sta="���ϳٵ�";
		    	else if(compare(time,m2)>0&&compare(time,n1)<0&&!closer(m2,time,n1))
		    		sta="��������";
		    	else if(compare(time,n1)>=0&&compare(time,n2)<=0)
		    		sta="��������";
		    	else if(compare(time,n2)>0&&compare(time,a)<0&&closer(n2,time,a))
		    		sta="����ٵ�";
		    	else if(compare(time,n2)>0&&compare(time,a)<0&&!closer(n2,time,a))
		    		sta="��������";
		    	else if(compare(time,a)>=0)
		    		sta="��������";
		    	Statement state2=conn.createStatement();
		    	count+=state2.executeUpdate("update kq set ״̬='"+sta+"' where ID="+id);
		    	state2.close();
		    }
		    rs.close();
		    state.close();
		    conn.close();
		    System.out.println(count);
		    
		} catch (SQLException e) {
		    System.out.println("SQL Exception: "+ e.toString());
		} catch (ClassNotFoundException cE) {
		    System.out.println("Class Not Found Exception: "+ cE.toString());
		}
		
		return count;
	}
	
	public int delete()
	{
		Connection conn = null;
		try {
		    Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
		    String url = "jdbc:odbc:Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=./db/cs.accdb";
		    conn = DriverManager.getConnection(url);
		    Statement state=conn.createStatement();
		    state.execute("delete from kq");
		    state.close();
		    conn.close();
		    return 1;
		} catch (SQLException e) {
		    System.out.println("SQL Exception: "+ e.toString());
		} catch (ClassNotFoundException cE) {
		    System.out.println("Class Not Found Exception: "+ cE.toString());
		}
		return 0;
	}
}
