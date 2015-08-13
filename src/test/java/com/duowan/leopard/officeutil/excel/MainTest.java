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

import com.duowan.leopard.officeutil.excel.bean.ExportExcelBean;

/**
 * @author lizeyang
 * @date 2015年6月3日
 */
public class MainTest extends TestCase {

	public void testExportByAnnotation() throws IllegalArgumentException, IllegalAccessException,
			FileNotFoundException, WriteException {

		List<TestBean> beanList = new ArrayList<TestBean>();

		for (int i = 0; i < 3000; i++) {
			TestBean bean = new TestBean();
			bean.setIntTest(10000000);
			bean.setStrTest("努力造轮子");
			bean.setTimeTest(new Timestamp(System.currentTimeMillis()));
			beanList.add(bean);
		}

		long now = System.currentTimeMillis();

		ExportExcelUtil util = new ExportExcelUtil();
		OutputStream out = new FileOutputStream("d:/yy-export-excel/test22.xls");
		util.exportByAnnotation(out, beanList);

		System.out.println(System.currentTimeMillis() - now);

	}

	public void testExport() throws WriteException {
		long start = System.currentTimeMillis();

		// 这个list,表示excel中的每一行数据
		List<Object> li = new ArrayList<Object>();

		// 在本例子中，每一行的数据，填充的是每个TestBean类的数据
		TestBean testBean = new TestBean();
		// 每一行有三列，分别是IntTest、StrTest，TimetTest
		testBean.setIntTest(8888);
		testBean.setStrTest("88888.888");
		testBean.setTimeTest(new Timestamp(System.currentTimeMillis()));
		for (int i = 0; i < 1000; i++) {
			li.add(testBean);
		}

		// 这里定义了列的先后顺序。按照put顺序不同，第一列是填充timeTest属性，列的标题显示为time类型。依次类推是第二列、第三列
		LinkedHashMap<String, String> keyMap = new LinkedHashMap<String, String>();
		keyMap.put("timeTest", "time类型");
		keyMap.put("intTest", "int类型");
		keyMap.put("strTest", "string类型");

		// 可以插入两个子表。并定义两个表的名字
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

		File file = new File("d:/yy-export-excel");
		file.mkdirs();

		// 运行一下，查看输出的excel
		try {
			ExportExcelUtil util = new ExportExcelUtil();
			OutputStream out = new FileOutputStream("d:/yy-export-excel/test.xls");
			util.export(sheetContentList, out);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}

		long end = System.currentTimeMillis();
		System.out.println(end - start);
	}
}
