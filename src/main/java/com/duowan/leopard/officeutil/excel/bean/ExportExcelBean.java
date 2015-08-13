package com.duowan.leopard.officeutil.excel.bean;

import java.util.LinkedHashMap;
import java.util.List;

/**
 * @author lizeyang
 * @date 2015年5月29日
 */
public class ExportExcelBean {
	/**
	 * 要填充的内容
	 */
	private List<Object> contentList;

	/**
	 * 表列标题名称
	 */
	private LinkedHashMap<String, String> keyMap;
	/**
	 * 分表名
	 */
	private String sheetName;

	public List<Object> getContentList() {
		return contentList;
	}

	public void setContentList(List<Object> contentList) {
		this.contentList = contentList;
	}

	public LinkedHashMap<String, String> getKeyMap() {
		return keyMap;
	}

	public void setKeyMap(LinkedHashMap<String, String> keyMap) {
		this.keyMap = keyMap;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

}
