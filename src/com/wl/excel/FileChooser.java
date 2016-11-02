package com.wl.excel;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map.Entry;

public class FileChooser extends JFrame implements ActionListener {
	private static final long serialVersionUID = 1L;
	JButton open = null;
	JPanel panelContainer = null;

	public FileChooser() {
		open = new JButton("浏览文件");
		open.setBounds(0, 0, 100, 100);
		this.setBounds(0, 0, 1280, 720);
		this.setVisible(true);
		this.setTitle("Excel 处理");
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		open.addActionListener(this);
		panelContainer = new JPanel();
		// panelContainer 的布局为 GridBagLayout
		panelContainer.setLayout(new GridBagLayout());
		GridBagConstraints c1 = new GridBagConstraints();
		c1.gridx = 0;
		c1.gridy = 0;
		c1.gridheight = 140;
		c1.weightx = 1.0;
		c1.weighty = 1.0;
		c1.fill = GridBagConstraints.NONE;
		panelContainer.add(open, c1);
		this.setContentPane(panelContainer);
        this.revalidate();
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		JFileChooser jfc = new JFileChooser();
		jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
		jfc.showDialog(new JLabel(), "选择");
		File file = jfc.getSelectedFile();
		
		if(file == null) {
			//TODO 
			System.out.println("文件不存在");
			return ;
		}
		if (file.isDirectory()) {
			System.out.println("文件夹:" + file.getAbsolutePath());
			System.out.println("不支持文件夹");
			return;
		} else if (file.isFile()) {
			System.out.println("文件:" + file.getAbsolutePath());		
		}
		
		System.out.println(jfc.getSelectedFile().getName());
		if (jfc.getSelectedFile().getName().endsWith(".xlsx")) {
			ExcelReader reader = new ExcelReader();
			List<Commodity> list = reader.readFromExcel(file.getAbsolutePath());
			if (list == null)
				return;
			GridBagConstraints c1 = new GridBagConstraints();
			c1.gridx = 0;
			c1.gridy = 1;
			c1.weightx = 1.0;
			c1.weighty = 1.0;
			c1.fill = GridBagConstraints.BOTH;
			panelContainer.add(createTopPanel(list), c1);
			
			GridBagConstraints c2 = new GridBagConstraints();
			c2.gridx = 0;
			c2.gridy = 2;
			c2.weightx = 0;
			c2.weighty = 0;
			c2.fill = GridBagConstraints.NONE;
			panelContainer.add(createTjPanel(list), c2);
            panelContainer.remove(0);
			panelContainer.revalidate();
		}
	}
	
	
	private JPanel createTjPanel(List<Commodity> list) {
		JPanel panel = new JPanel();
		panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
		
		GridBagConstraints c = new GridBagConstraints();
		c.gridx = 0;
		c.gridy = 1;
		c.weightx = 0;
		c.weighty = 0;
		c.fill = GridBagConstraints.NONE;
		int macCount = ExcelReader.getCountByType(Commodity.TYPE_MAC, list);
		int macSum = ExcelReader.getSumByType(Commodity.TYPE_MAC, list);
		JPanel macItemPanel = createItem("电脑"+ " : ", macCount + "台 总额:" + macSum);
		panel.add(macItemPanel, c);
		
		int padCount = ExcelReader.getCountByType(Commodity.TYPE_PAD, list);
		int padSum = ExcelReader.getSumByType(Commodity.TYPE_PAD, list);
		JPanel padItemPanel = createItem("平板"+ " : ", padCount + "台 总额:" + padSum);
		panel.add(padItemPanel, c);
		
		int phoneCount = ExcelReader.getCountByType(Commodity.TYPE_PHONE, list);
		int phoneSum = ExcelReader.getSumByType(Commodity.TYPE_PHONE, list);
		JPanel phoneItemPanel = createItem("手机"+ " : ", phoneCount + "台 总额:" + phoneSum);
		panel.add(phoneItemPanel, c);
		
		int dlCount = ExcelReader.getCountByType(Commodity.TYPE_DATALINE, list);
		int dlSum = ExcelReader.getSumByType(Commodity.TYPE_DATALINE, list);
		JPanel dlItemPanel = createItem("数据线"+ " : ", dlCount + "条 总额:" + dlSum);
		panel.add(dlItemPanel, c);
		
		JPanel didvertemPanel = createItem("按人统计", "");
		panel.add(didvertemPanel, c);
		
		HashMap<String, List<Commodity>> map = ExcelReader.groupByPerson(list);
		Iterator<Entry<String, List<Commodity>>> itr = map.entrySet().iterator();
		while (itr.hasNext()) {
			Entry<String, List<Commodity>> entry = (Entry<String, List<Commodity>>) itr.next();
			List<Commodity> subList = entry.getValue();
			StringBuilder builder = new StringBuilder();
			builder.append("电脑：").append(ExcelReader.getCountByType(Commodity.TYPE_MAC, subList)).append("台").append(",");
			builder.append("平板：").append(ExcelReader.getCountByType(Commodity.TYPE_PAD, subList)).append("台").append(",");
			builder.append("手机：").append(ExcelReader.getCountByType(Commodity.TYPE_PHONE, subList)).append("台").append(",");
			builder.append("数据线：").append(ExcelReader.getCountByType(Commodity.TYPE_DATALINE, subList)).append("条").append(",");
			builder.append("配件数：").append(ExcelReader.getComponentCount(subList)).append("个").append(",");
			builder.append("配件总额：").append(ExcelReader.getComponentSum(subList)).append("").append(",");
			builder.append(" 提成：").append(ExcelReader.getPercentage(subList));
			builder.append(" 合计总额：").append(ExcelReader.getSum(subList));
			JPanel itemPanel = createItem(entry.getKey() + "   ", builder.toString());
			panel.add(itemPanel, c);
		}
		
		panel.add(createItem(" ", ""), c);
		panel.add(createItem("合计", "配件数：" + ExcelReader.getComponentCount(list) + ", 配件总额：" + ExcelReader.getComponentSum(list) + ", 总额：" + ExcelReader.getSum(list)), c);
		
		return panel;
	}
	
	private JPanel createItem(String title, String content) {
		JPanel itemPanel = new JPanel();
		itemPanel.setLayout(new BoxLayout(itemPanel, BoxLayout.X_AXIS));
		JLabel tLable = new JLabel(title);
		GridBagConstraints c = new GridBagConstraints();
		c.gridx = 0;
		c.gridy = 1;
		c.weightx = 0;
		c.weighty = 0;
		c.fill = GridBagConstraints.VERTICAL;
		itemPanel.add(tLable, c);

		JLabel cLable = new JLabel(content);
		itemPanel.add(cLable, c);
		return itemPanel;
	}
	
	
	private String[][] convertToArr(List<Commodity> list) {
		if(list == null)
			return null;
		String[][] arr = new String[list.size()][8];
		if(list != null) {
			for (int i = 0; i < list.size(); i++) {
				composeData(arr[i], list.get(i));
			}
		}
	
		return arr;
	}
	
	private void composeData(String[] arr, Commodity commodity) {
		if(arr.length < 8)
			return ;
		arr[0] = commodity.getDatetime();
		arr[1] = commodity.getName();
		arr[2] = commodity.getTypeName();
		arr[3] = String.valueOf(commodity.getPrice());
		arr[4] = String.valueOf(commodity.getCount());
		arr[5] = commodity.getAssistant();
		arr[6] = String.valueOf(commodity.getPercentage());
		arr[7] = "";
	}
	
	private JPanel createTopPanel(List<Commodity> list) {
		JPanel topPanel = new JPanel();
		String[] columnName = { "日期", "物品名称", "类型", "价格","数量","售卖人","提成", "备注" };
		
		String[][] rowData = convertToArr(list);
			
		// 创建表格
		JTable table = new JTable(new DefaultTableModel(rowData, columnName));
		// 创建包含表格的滚动窗格
		JScrollPane scrollPane = new JScrollPane(table);
		scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
		// 定义 topPanel 的布局为 BoxLayout，BoxLayout 为垂直排列
		topPanel.setLayout(new BoxLayout(topPanel, BoxLayout.Y_AXIS));
		// 先加入一个不可见的 Strut，从而使 topPanel 对顶部留出一定的空间
		topPanel.add(Box.createVerticalStrut(10));
		// 加入包含表格的滚动窗格
		topPanel.add(scrollPane);
		// 再加入一个不可见的 Strut，从而使 topPanel 和 middlePanel 之间留出一定的空间
		topPanel.add(Box.createVerticalStrut(10));
		return topPanel;
	}

}
