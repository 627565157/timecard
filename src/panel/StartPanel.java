package panel;

import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;

import readExcel.PoiReadExcel;

public class StartPanel {

	private JFrame frame;
	private JTextField execlUrl;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					StartPanel window = new StartPanel();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public StartPanel() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 443, 419);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);

		JLabel lblExcel = new JLabel("Excel\u6587\u4EF6\u8DEF\u5F84\uFF1A");
		lblExcel.setBounds(37, 53, 97, 24);
		frame.getContentPane().add(lblExcel);

		execlUrl = new JTextField();
		execlUrl.setBounds(144, 55, 210, 21);
		frame.getContentPane().add(execlUrl);
		execlUrl.setColumns(10);

		JButton button = new JButton("\u5BFC\u51FA\u7ED3\u679C");
		button.addActionListener(new ActionListener() {
			@SuppressWarnings("static-access")
			public void actionPerformed(ActionEvent e) {
				long start = System.currentTimeMillis();
				PoiReadExcel.calculateResult(execlUrl.getText().toString().replace('\\', '/'));
				// DBUtil db = new DBUtil();
				// Read read = new Read();
				// Export ex = new Export();
				// try {
				// String path =
				// execlUrl.getText().toString().replace('\\','/');
				// db.insert(read.getAll(path));
				// Export.date = read.getAll(path).get(0).get(2);
				// db.update();
				// ex.export();
				// } catch (SQLException e1) {
				// // TODO Auto-generated catch block
				// e1.printStackTrace();
				// }
				// db.delete();
				long end = System.currentTimeMillis();
				System.out.println("运行时间：" + (end - start) + "毫秒");
				JOptionPane.showMessageDialog(null, "导出结果成功！", "结果", JOptionPane.DEFAULT_OPTION);
			}
		});
		button.setBounds(119, 240, 153, 23);
		frame.getContentPane().add(button);
	}
}
