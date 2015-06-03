package com.duowan.leopard.officeutil.excel;

import java.util.LinkedHashMap;
import java.util.List;

/**
 * @author lizeyang
 * @date 2015年5月29日
 */
public class ExportExcelBean {
	private List<Object> contentList;

	private LinkedHashMap<String, String> keyMap;

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
