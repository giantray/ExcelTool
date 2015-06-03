package com.duowan.leopard.officeutil.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

import junit.framework.TestCase;
import jxl.write.WriteException;

/**
 * @author lizeyang
 * @date 2015年6月3日
 */
public class Test extends TestCase {

	public void testExport() throws WriteException {
		long start = System.currentTimeMillis();

		List<Object> li = new ArrayList<Object>();

		TestBean testBean = new TestBean();
		testBean.setIntTest(8888);
		testBean.setStrTest("88888.888");
		testBean.setTimeTest(new Timestamp(System.currentTimeMillis()));

		for (int i = 0; i < 1000; i++) {
			li.add(testBean);
		}

		LinkedHashMap<String, String> keyMap = new LinkedHashMap<String, String>();
		keyMap.put("timeTest", "time类型");
		keyMap.put("intTest", "int类型");
		keyMap.put("strTest", "string类型");

		List<ExportExcelBean> sheetContentList = new ArrayList<ExportExcelBean>();
		ExportExcelBean bean1 = new ExportExcelBean();
		bean1.setContentList(li);
		bean1.setKeyMap(keyMap);
		bean1.setSheetName("测试1");

		ExportExcelBean bean2 = new ExportExcelBean();
		bean2.setContentList(li);
		bean2.setKeyMap(keyMap);
		bean2.setSheetName("测试2");

		sheetContentList.add(bean1);

		sheetContentList.add(bean2);

		OutputStream out;
		try {
			ExportExcelUtil util = new ExportExcelUtil();
			File file = new File("d:/yy-export-excel");
			file.mkdirs();

			out = new FileOutputStream("d:/yy-export-excel/test.xls");
			util.export(sheetContentList, out);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}

		long end = System.currentTimeMillis();
		System.out.println(end - start);
	}
}
