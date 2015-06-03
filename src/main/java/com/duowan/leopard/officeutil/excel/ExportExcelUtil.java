package com.duowan.leopard.officeutil.excel;

/**
 * @author lizeyang
 * @date 2015年5月29日
 */

import java.io.OutputStream;
import java.lang.reflect.Field;
import java.sql.Timestamp;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import jxl.CellView;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExportExcelUtil {

	public static final int RESULT_SUCC = 0;
	public static final int RESULT_FAIL = -1;
	public static final String TYPE_YYYY_MM_DD_HH_MM_SS = "yyyy-MM-dd HH:mm:ss";
	public static final int fontSize = 10;

	// 两种字体样式，一种正常样式，一种是粗体样式
	WritableFont NormalFont = new WritableFont(WritableFont.ARIAL, fontSize);
	WritableFont BoldFont = new WritableFont(WritableFont.ARIAL, fontSize, WritableFont.BOLD);

	// 标题（列头）样式
	WritableCellFormat titleFormat = new WritableCellFormat(BoldFont);
	// 正文样式1：居中
	WritableCellFormat contentCenterFormat = new WritableCellFormat(NormalFont);
	// 正文杨师：右对齐
	WritableCellFormat contentRightFormat = new WritableCellFormat(NormalFont);

	WritableWorkbook workbook;

	public ExportExcelUtil() throws WriteException {

		titleFormat.setBorder(Border.ALL, BorderLineStyle.THIN); // 线条
		titleFormat.setVerticalAlignment(VerticalAlignment.CENTRE);

		titleFormat.setAlignment(Alignment.CENTRE); // 文字对齐
		titleFormat.setWrap(false); // 文字是否换行

		contentCenterFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		contentCenterFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
		contentCenterFormat.setAlignment(Alignment.CENTRE);
		contentCenterFormat.setWrap(false);

		contentRightFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		contentRightFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
		contentRightFormat.setAlignment(Alignment.RIGHT);
		contentRightFormat.setWrap(false);
	}

	/**
	 * 将数据转成成excel。 特性： 1、将时间类型的值转成yyyy-MM-dd HH:mm:ss 2、将数字类型的值转成带千分符的形式，并右对齐
	 * 3、除数字类型外，其他类型的值居中显示
	 *
	 * @param sheetContentList封装表格内容
	 *            ，其中参数keyMap、contentList含义，如下所示
	 * @param keyMap
	 *            定义标题及每一列对应的JavaBean属性。标题的先后顺序，对应keyMap的插入顺序；
	 *            map中的key值为JavaBean属性，value为标题
	 * @param contentList
	 *            表格内容，List中的每一个元素，对应到excel的每一行
	 * @param os
	 *            结果输出流
	 * @return
	 */
	public final int export(List<ExportExcelBean> sheetContentList, OutputStream os) {

		int rs = RESULT_SUCC;
		try {
			workbook = Workbook.createWorkbook(os);

			for (int i = 0; i < sheetContentList.size(); i++) {
				addSheet(sheetContentList.get(i).getKeyMap(), sheetContentList.get(i).getContentList(),
						sheetContentList.get(i).getSheetName(), i);
			}

			workbook.write();
			workbook.close();
		} catch (Exception e) {
			rs = RESULT_FAIL;
			e.printStackTrace();
		}
		return rs;
	}

	public final int export(LinkedHashMap<String, String> keyMap, List<Object> contentList, OutputStream os) {
		List<ExportExcelBean> list = new ArrayList<ExportExcelBean>();
		ExportExcelBean bean = new ExportExcelBean();
		bean.setContentList(contentList);
		bean.setKeyMap(keyMap);
		bean.setSheetName("sheet1");
		list.add(bean);

		return export(list, os);
	}

	private void addSheet(LinkedHashMap<String, String> keyMap, List<Object> listContent, String sheetName, int sheetNum)
			throws WriteException, RowsExceededException, NoSuchFieldException, IllegalAccessException {
		// 创建名为sheetName的工作表
		WritableSheet sheet = workbook.createSheet(sheetName, sheetNum);

		// 设置标题,标题内容为keyMap中的value值,标题居中粗体显示
		Iterator titleIter = keyMap.entrySet().iterator();
		int titleIndex = 0;
		while (titleIter.hasNext()) {
			Map.Entry<String, String> entry = (Map.Entry<String, String>) titleIter.next();
			sheet.addCell(new Label(titleIndex++, 0, entry.getValue(), titleFormat));
		}

		// 设置正文内容
		for (int i = 0; i < listContent.size(); i++) {
			Iterator contentIter = keyMap.entrySet().iterator();
			int colIndex = 0;
			int listIndex = 0;
			while (contentIter.hasNext()) {
				Map.Entry<String, String> entry = (Map.Entry<String, String>) contentIter.next();
				Object key = entry.getKey();

				Field field = listContent.get(i).getClass().getDeclaredField(key.toString());
				field.setAccessible(true);
				Object content = field.get(listContent.get(i));

				String contentStr = null != content ? content.toString() : "";

				WritableCellFormat cellFormat = contentCenterFormat;

				// 将数字转变成千分位格式
				String numberStr = getNumbericValue(contentStr);
				// numberStr不为空，说明是数字类型。
				if (null != numberStr && !numberStr.trim().equals("")) {
					contentStr = numberStr;
					// 数字要右对齐
					cellFormat = contentRightFormat;
				} else {
					// 如果是时间类型。要格式化成标准时间格式
					String timeStr = getTimeFormatValue(field, content);
					// timeStr不为空，说明是时间类型
					if (null != timeStr && !timeStr.trim().equals("")) {
						contentStr = timeStr;
					}
				}

				sheet.addCell(new Label(colIndex++, i + 1, contentStr, cellFormat));

			}

		}

		setAutoSize(sheet, keyMap.size(), listContent.size());
	};

	/**
	 * 宽度自适应
	 * 
	 * @param sheet
	 * @param colNum
	 * @param rowNum
	 * @return
	 */
	private boolean setAutoSize(WritableSheet sheet, int colNum, int rowNum) {

		for (int i = 0; i < colNum; i++) {

			int maxLength = 0;

			CellView cell = sheet.getColumnView(i);

			for (int j = 0; j < rowNum; j++) {
				maxLength = Math.max(sheet.getCell(i, j).getContents().getBytes().length, maxLength);
			}

			cell.setSize(25 * fontSize * maxLength);

			sheet.setColumnView(i, cell);
		}

		return true;
	}

	/**
	 * 获取格式化后的时间串
	 *
	 * @param field
	 * @param content
	 * @return
	 */
	private String getTimeFormatValue(Field field, Object content) {
		String timeFormatVal = "";
		if (field.getType().getName().equals(java.sql.Timestamp.class.getName())) {
			Timestamp time = (Timestamp) content;
			timeFormatVal = longTimeTypeToStr(time.getTime(), TYPE_YYYY_MM_DD_HH_MM_SS);
		} else if (field.getType().getName().equals(java.util.Date.class.getName())) {
			Date time = (Date) content;
			timeFormatVal = longTimeTypeToStr(time.getTime(), TYPE_YYYY_MM_DD_HH_MM_SS);
		}

		return timeFormatVal;
	}

	/**
	 * 获取千分位数字
	 *
	 * @param str
	 * @return
	 */
	private String getNumbericValue(String str) {
		String numbericVal = "";
		try {
			Double doubleVal = Double.valueOf(str);
			numbericVal = DecimalFormat.getNumberInstance().format(doubleVal);
		} catch (NumberFormatException e) {
			// if exception, not format
		}
		return numbericVal;
	}

	/**
	 * 格式化时间
	 *
	 * @param time
	 * @param formatType
	 * @return
	 */
	public String longTimeTypeToStr(long time, String formatType) {

		String strTime = "";
		if (time >= 0) {
			SimpleDateFormat sDateFormat = new SimpleDateFormat(formatType);

			strTime = sDateFormat.format(new Date(time));

		}

		return strTime;

	}

	public static void main(String[] args) {

	}

}