package com.duowan.leopard.officeutil.excel;

import java.sql.Timestamp;

import com.duowan.leopard.officeutil.excel.annotation.ExcelSheet;
import com.duowan.leopard.officeutil.excel.annotation.SheetCol;

/**
 * @author lizeyang
 * @date 2015年6月3日
 */
@ExcelSheet(name = "这是表的名字", order = "strTest")
public class TestBean {
	@SheetCol("字符串")
	private String strTest;
	@SheetCol("数字")
	private int intTest;
	@SheetCol("时间")
	private Timestamp timeTest;

	public String getStrTest() {
		return strTest;
	}

	public void setStrTest(String strTest) {
		this.strTest = strTest;
	}

	public int getIntTest() {
		return intTest;
	}

	public void setIntTest(int intTest) {
		this.intTest = intTest;
	}

	public Timestamp getTimeTest() {
		return timeTest;
	}

	public void setTimeTest(Timestamp timeTest) {
		this.timeTest = timeTest;
	}
}
