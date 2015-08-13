package com.duowan.leopard.officeutil.excel;

import java.lang.reflect.Field;

import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WriteException;

/**
 * @author lizeyang
 * @date 2015年8月13日
 */
public class ExtendUtil extends ExportExcelUtil {
	public ExtendUtil() throws WriteException {
		super();
	}

	/**
	 * 说明，你可以继承ExportExcelUtil,然后改写方法，以实现自定义的目的。例如这里演示了改写getContentLabel方法，
	 * 以改变每个格子中填写内容的目录
	 */
	protected Label getContentLabel(int col, int row, Field field, Object content) {
		WritableCellFormat cellFormat = contentCenterFormat;
		String contentStr = "change";
		Label label = new Label(col, row, contentStr, cellFormat);
		return label;
	};
}
