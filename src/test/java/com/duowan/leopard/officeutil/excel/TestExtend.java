package com.duowan.leopard.officeutil.excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.List;

import jxl.write.WriteException;

import org.junit.Test;

/**
 * @author lizeyang
 * @date 2015年8月13日
 */

public class TestExtend {

	@Test
	public void test() throws WriteException, FileNotFoundException {
		ExtendUtil extendUtil = new ExtendUtil();
		List<TestBean> beanList = new ArrayList<TestBean>();

		for (int i = 0; i < 3000; i++) {
			TestBean bean = new TestBean();
			bean.setIntTest(10000000);
			bean.setStrTest("努力造轮子");
			bean.setTimeTest(new Timestamp(System.currentTimeMillis()));
			beanList.add(bean);
		}

		OutputStream out = new FileOutputStream("d:/yy-export-excel/test3.xls");
		extendUtil.exportByAnnotation(out, beanList);
	}

}
