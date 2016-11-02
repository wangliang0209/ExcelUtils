package com.wl.excel;

/**
 * 商品信息
 * @author wangliang
 *
 */
public class Commodity {
	public static final int TYPE_MAC = 1;//电脑
	public static final int TYPE_PAD = 2;//平板
	public static final int TYPE_PHONE = 3;//手机
	public static final int TYPE_DATALINE = 4;//数据线
	public static final int TYPE_COMPONENT = 5; //配件
	private String datetime;
	private int type;
	private String name;
	private double price;
	private int count;
	private String assistant;//售货员
	private double percentage;//提成
	
	public int getType() {
		return type;
	}
	public void setType(int type) {
		this.type = type;
	}
	public String getDatetime() {
		return datetime;
	}
	public void setDatetime(String datetime) {
		this.datetime = datetime;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public double getPrice() {
		return price;
	}
	public void setPrice(double price) {
		this.price = price;
	}
	public int getCount() {
		return count;
	}
	public void setCount(int count) {
		this.count = count;
	}
	public String getAssistant() {
		return assistant;
	}
	public void setAssistant(String assistant) {
		this.assistant = assistant;
	}
	public double getPercentage() {
		return percentage;
	}
	public void setPercentage(double percentage) {
		this.percentage = percentage;
	}
	
	public String getTypeName(){
		String str = null;
		switch (type) {
		case TYPE_MAC:
			str = "笔记本电脑";
			break;
		case TYPE_PAD:
			str = "平板";
			break;
		case TYPE_PHONE:
			str = "手机";
			break;
		case TYPE_DATALINE:
			str = "数据线";
			break;
		default:
			str = "其他";
			break;
		}
		return str;
	}
	
	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder();
		builder.append("datetime:").append(datetime).append(",")
			.append("type:").append(type).append(",")
			.append("name:").append(name).append(",")
			.append("price:").append(price).append(",")
			.append("count:").append(count).append(",")
			.append("assistant:").append(assistant).append(",")
			.append("percentage:").append(percentage);
		return builder.toString();
	}
	
}
