##概述
通过对Jxl包的简单封装，这个组件提供了一个更加灵活、简单地生成excel表格的方法。
在生成excel时，你只需编码，把数据塞到指定的bean中，之后生成excel的其他细节，交给这个工具。这个工具已经对excel样式做了较佳的设置，可以省却你调试excel样式的痛苦

##特性
- **解耦** 在jxl包基础上做了小扩展，程序员只要专注塞数据——把数据塞到我们提供的bean。之后bean转excel，由这个组件搞定。该组件已经优化了表格样式，可以省却调试excel样式的痛苦
- **灵活排序** 可以很方便地调整每列的先后顺序
- **宽度自适应** 针对中文做优化，确保excel表格能依据文字宽度确定每个表格宽度
- **值类型自动识别** 检测每个表格的内容，是数字、字符串、还是时间，依据类型的不同，做人性化、合理的显示

##引入
maven依赖

```xml
<dependency>
  <groupId>com.github.giantray</groupId>
  <artifactId>exceltool</artifactId>
  <version>0.95</version>
</dependency>
```
当前版本为0.95


##如何使用
提供了两种使用方式，一种是基于传参的方式，一种是基于注解的方式。

###2、基于传参
所谓基于参数，是指excel中要展示的列，列名，顺序等通过构建一个参数bean来指定。

2.1、在MainTest类中，提供了一个测试类，为你演示了如何生成，建议您直接看看测试类中的demo。下面介绍这个demo

2.2、首先，demo代码中，会将数据塞到ExportExcelBean这个类中
```java
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
```
如上所示演示了将一些测试数据填入到ExportExcelBean中。ExportExcelBean有三个属性
sheetname:表名。在一个excel文件中，允许存在多个表。每个表的名字，一般会在excel左下角显示，如下所示
![上传图片](http://image.game.yy.com/o/cloudapp/25586759/170x170/201506-bbc2a60f_094e_498b_87e7_2ead79ca9536.png)

contentList:要填充的内容。表的每一行，是一个对象，多个对象一起，就组成了这个contentList。对象的类，可以依据你自己的实际情况，用自己写的类。本文例子中，是将数据填充到TestBean 这个类中

keyMap:表列名及属性映射关系。请注意，这是一个LinkedHashMap，也就是说，是有顺序先后性的。之后你要调整每一列的顺序，也只要调整这里的顺序即可
如例子中这三行代码，表示表的第一列，列的标题为time类型，填充的值为TestBean 类中timeTest这个属性。第二列标题则为“int类型”，填充TestBean 中intTest属性
```java
		keyMap.put("timeTest", "time类型");
		keyMap.put("intTest", "int类型");
		keyMap.put("strTest", "string类型");

```


###3、基于注解
基于注解的方式，指的是excel的列，列名，顺序等由源数据bean中的注解来指定。测试类TestExtend演示了这个方式

```java
@ExcelSheet(name = "这是表的名字", order = "strTest,intTest")
public class TestBean {
	@SheetCol("字符串")
	private String strTest;
	@SheetCol("数字")
	private int intTest;
	@SheetCol("时间")
	private Timestamp timeTest;

	//这里省略get、set方法
｝

```

这里假设我们要将List<TestBean>的数据输出为excel，可以像上面这样做注解。
- 注解ExcelSheet，name属性指定了表的名字，order属性指定了列的顺序，其中的值是bean中的属性名，多个属性用逗号隔开
- 注解SheetCol,带有这个注解的属性，才会被输出到excel中，且表头的列名，为这里指定的名字

之后调用exportByAnnotation方法，就可以得到excel文件