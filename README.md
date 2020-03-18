# 简介

Apache POI是一套基于 OOXML 标准（Office Open XML）和 OLE2 标准来读写各种格式文件的 Java API，也就是说只要是遵循以上标准的文件，POI 都能够进行读写，而不仅仅只能操作我们熟知的办公程序文件。本文只会涉及到 excel 相关内容，其他文件的操作可以参考[poi官方网站]( https://poi.apache.org/components )。

这里先总结下 POI 的使用体验。POI 面向接口的设计非常巧妙，使用 `ss.usermodel` 包读写 xls 和 xlsx 时，可以使用同一套代码，即使这两种文件格式采用的是完全不同的标准。POI 提供了`SXSSFWorkbook`用于解决 xlsx 写大文件时容易出现的 OOM 问题。但是，还是存在以下不足（都只针对读场景）：

1. **使用 ss.usermodel 包解析 excel 时比较占内存，容易出现 OOM**。类似于 xml 中的 DOM，这种方式会在内存中构建整个文档的结构，在处理大文件时容易出现 OOM。然而，**大部分场景我们并不需要随机地去访问 excel 中的节点**。
2. **POI 提供的 SAX API 太过复杂**。为了解决第一个问题，POI 提供了基于事件驱动的 SAX 方式，这种方式内存占用小、效率高， 但是 API 太过繁琐，开发者必须在熟知文档规范的前提下才能使用，而且 xls 和 xlsx 使用的是完全不同的两套 API，实际项目中必须针对不同文件类型分别实现。这一点可以从本文的例子看出来。

针对以上问题，阿里的 easyexcel 对 POI 进行高级封装，简化了 SAX 部分 API 的使用，读 excel 时只采用 SAX 方式，从而避免出现 OOM，并且重写了 POI 对 xlsx 的解析，能够原本一个3M的 excel 用 POI SAX 依然需要100M左右内存降低到几M。

## 什么是 OLE2 和 OOXML

OLE2 和 OOXML 本质上都是一种文件格式规范或标准，平时看到的 excel 中，有字体、公式、颜色、图片等等，看起来非常复杂，但是在文件结构上都遵循着固定的格式。

OLE2 文件一般包括 xls、doc、ppt 等，是二进制格式的文件。 相关内容可以参考：[ 复合文档Ole对象二进制储存格式 ]( 复合文档Ole对象二进制储存格式 )。

OOXML文件一般包括 xlsx、docx、pptx 等。该类文件以指定格式的 xml 为基础并以 zip 格式压缩，这里我利用解压工具解压本地的一个 xml 文件，可以看到以下文件结构，在本文例子中，我们会重点关注 sharedStrings.xml 和 sheet1.xml 的内容，因为使用 SAX API 时必须用到：

<img src="D:\growUp\git_repository\09-poi-demo\extend\img\xlsx文件结构.png" alt="xlsx文件结构" style="zoom:80%;" />

## POI的组件

针对不同应用的文件，使用时需要引入对应的 maven 依赖，这里给出官方给出的指引。如果我们不使用 SAX API 方式读写 excel，一般只会用到这个 org.apache.poi.ss 中的 API，具体的实现类放在 org.apache.poi.hssf 或 org.apache.poi.xssf 。

|                   组件                    |              作用               |                          Maven依赖                           |
| :---------------------------------------: | :-----------------------------: | :----------------------------------------------------------: |
|                   POIFS                   |         OLE2 Filesystem         |                            *poi*                             |
|                   HPSF                    |       OLE2 Property Sets        |                            *poi*                             |
|   <font color='red'><b>HSSF</b></font>    |            Excel XLS            |                            *poi*                             |
|                   HSLF                    |         PowerPoint PPT          |                       *poi-scratchpad*                       |
|                   HWPF                    |            Word DOC             |                       *poi-scratchpad*                       |
|                   HDGF                    |            Visio VSD            |                       *poi-scratchpad*                       |
|                   HPBF                    |          Publisher PUB          |                       *poi-scratchpad*                       |
|                   HSMF                    |           Outlook MSG           |                       *poi-scratchpad*                       |
|                    DDF                    |     Escher common drawings      |                            *poi*                             |
|                   HWMF                    |          WMF drawings           |                       *poi-scratchpad*                       |
|                 OpenXML4J                 |              OOXML              | *poi-ooxml* plus either *poi-ooxml-schemas* or *ooxml-schemas* and *ooxml-security* |
|   <font color='red'><b>XSSF</b></font>    |           Excel XLSX            |                         *poi-ooxml*                          |
|                   XSLF                    |         PowerPoint PPTX         |                         *poi-ooxml*                          |
|                   XWPF                    |            Word DOCX            |                         *poi-ooxml*                          |
|                   XDGF                    |           Visio VSDX            |                         *poi-ooxml*                          |
|                 Common SL                 | PowerPoint PPT 和 PPTX 共用组件 |               *poi-scratchpad* and *poi-ooxml*               |
| <font color='red'><b>Common SS</b></font> |   Excel XLS 和 XLSX 共用组件    |                         *poi-ooxml*                          |

# 怎么使用POI

## 工程环境

JDK：1.8.0_201

maven：3.6.1

IDE：Spring Tool Suite  4.3.2.RELEASE

POI：4.1.2

easyexcel：2.1.6

mysql：5.7.28

## 创建项目

项目类型Maven Project，打包方式 jar。

## 引入依赖

```xml
		<!-- junit -->
		<dependency>
			<groupId>junit</groupId>
			<artifactId>junit</artifactId>
			<version>4.12</version>
			<scope>test</scope>
		</dependency>
		<!-- poi-ooxml -->
		<dependency>
		    <groupId>org.apache.poi</groupId>
		    <artifactId>poi-ooxml</artifactId>
		    <version>4.1.2</version>
		</dependency>
		<!-- easyexcel -->
		<dependency>
			<groupId>com.alibaba</groupId>
			<artifactId>easyexcel</artifactId>
			<version>2.1.6</version>
		</dependency>
        <!-- hikari -->
        <dependency>
            <groupId>com.zaxxer</groupId>
            <artifactId>HikariCP</artifactId>
            <version>2.6.1</version>
        </dependency>
        <!-- mysql驱动 -->
        <dependency>
            <groupId>mysql</groupId>
            <artifactId>mysql-connector-java</artifactId>
            <version>8.0.15</version>
        </dependency>
```

## 读excel--入门案例

### 需求

读取指定 excel 第一个单元格的内容。该指定文件第一个单元格内容为“测试”。

### 编写测试方法

建议采用`WorkbookFactory`来获取`Workbook`实例，而不是根据文件类型写死具体的实现类。

```java
	@Test
	public void test01() throws IOException {
		// 处理XSSF
		String path = "extend\\file\\poi_test_01.xlsx";
		// 处理HSSF
		//String path = "extend\\file\\poi_test_01.xls";
		
		// 创建工作簿，会根据excel命名选择不同的Workbook实现类
		Workbook wb = WorkbookFactory.create(new File(path));

		// 获取工作表
		Sheet sheet = wb.getSheetAt(0);

		// 获取行
		Row row = sheet.getRow(0);

		// 获取单元格
		Cell cell = row.getCell(0);

		// 获取单元格内容
		String value = cell.getStringCellValue();
		System.err.println("第一个单元格字符：" + value);

		// 释放资源
		wb.close();
	}
```

### 测试

运行以上方法，控制台打印出第一个单元格的内容：

<img src="D:\growUp\git_repository\09-poi-demo\extend\img\ReadTest01.png" alt='ReadTest01' style="zoom: 80%;" />

## 写excel--入门案例

### 需求

生成一个 excel 文件，并给第一个单元格赋值为"测试"。

### 编写测试方法

`CellUtil`是 POI 自带的工具类，这里简化了三句代码（创建单元格，设置样式，赋值）。注意，当写入 xlsx 的大文件时，可以考虑使用`SXSSFWorkbook`来避免 OOM。

```java
	@Test
	public void test01() throws FileNotFoundException, IOException {
		// 处理XSSF
		String path = "extend\\file\\poi_test_01.xlsx";
		// 处理HSSF
		// String path = "extend\\file\\poi_test_01.xls";

		// 创建工作簿
		boolean flag = path.endsWith(".xlsx");
		Workbook wb = WorkbookFactory.create(flag ? true : false);
		// Workbook wb = new SXSSFWorkbook(100);//内存仅保留100行数据，可避免OOM

		// 创建工作表
		Sheet sheet = wb.createSheet(WorkbookUtil.createSafeSheetName("MySheet001"));
		// 设置列宽
		sheet.setColumnWidth(0, 26 * 256);

		// 创建行(索引从0开始)
		Row row = sheet.createRow(0);
		// 设置行高
		row.setHeightInPoints(20.25f);

		// 创建单元格样式对象
		CellStyle style = wb.createCellStyle();
		// 设置样式
		style.setAlignment(HorizontalAlignment.CENTER); // 横向居中
		style.setVerticalAlignment(VerticalAlignment.CENTER);// 纵向居中
		style.setBorderBottom(BorderStyle.THIN);
		style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		// 创建单元格、设置样式和内容
		CellUtil.createCell(row, 0, "测试", style);

		// 保存到本地目录
		OutputStream out = new FileOutputStream(new File(path));
		wb.write(out);

		// 释放资源
		out.close();
		wb.close();
	}
```

### 测试

运行以上方法，指定路径下生成了 excel 文件，并填充了第一个单元格：

<img src="D:\growUp\git_repository\09-poi-demo\extend\img\WriteTest01.png" alt="WriteTest01" style="zoom: 67%;" />

## 读excel--批量导入excel数据到数据库

### 需求

将 excel 中的用户数据导入到数据库（sql 已提供，在当前项目的 extend/sql 下），数据格式如下：

<img src="D:\growUp\git_repository\09-poi-demo\extend\img\ReadTest02.png" alt="ReadTest02" style="zoom:67%;" />

该文件总计1000条数据，xls 大小 128 KB，xlsx 大小 40 KB，两种类型文件内容一致。

### 编写测试方法

一般 excel 的内容格式是提前约定好的，我们知道用户数据哪一列是用户名，哪一列是电话号码，所以，在获取单元格数据后可以准确地转换，但这种方式需要针对不同的对象分别定义一个转换方法。

```java
	@Test
	public void test02() throws SQLException, IOException {
		// 处理XSSF
		String path = "extend\\file\\user_data.xlsx";
		// 处理HSSF
		//String path = "extend\\file\\user_data.xls";
		
		// 定义集合，用于存放excel中的用户数据
		List<UserDTO> list = new ArrayList<>();

		InputStream in = new FileInputStream(path);
		// 创建工作簿
		Workbook wb = WorkbookFactory.create(in);

		// 获取工作表
		Sheet sheet = wb.getSheetAt(0);

		// 获取所有行
		Iterator<Row> iterator = sheet.iterator();

		int rowNum = 0;
		// 遍历行
		while(iterator.hasNext()) {
			Row row = iterator.next();
			// 跳过标题行
			if(rowNum == 0 || rowNum == 1) {
				rowNum++;
				continue;
			}
			// 将用户对象保存到集合中
			list.add(constructUserByRow(row));
		}
		// 批量保存
		new UserService().save(list);

		// 释放资源
		in.close();
		wb.close();
	}
	/**
	 * <p>通过行数据构造用户对象</p>
	 */
	private UserDTO constructUserByRow(Row row) {
		UserDTO userDTO = new UserDTO();
		Cell cell = null;
		// 用户名
		cell = row.getCell(1);
		userDTO.setName(cell.getStringCellValue());
		// 性别
		cell = row.getCell(2);
		userDTO.setGenderStr(cell.getStringCellValue());
		// 年龄
		cell = row.getCell(3);
		userDTO.setAge(((Double)cell.getNumericCellValue()).intValue());
		// 电话
		cell = row.getCell(4);
		userDTO.setPhone(cell.getStringCellValue());

		return userDTO;
	}
```

### 测试

运行以上方法，可以在数据库看到导入的数据：

<img src="D:\growUp\git_repository\09-poi-demo\extend\img\ReadTest03.png" alt="ReadTest03" style="zoom:80%;" />

## 写excel--批量导出数据库数据到excel

### 需求

将数据库的用户数据导出到excel中。这个例子使用模板进行导出，模板如下（如果是 xlsx 的大文件，为了能够使用`SXSSFWorkbook`最好不要用模板）。

<img src="D:\growUp\git_repository\09-poi-demo\extend\img\WriteTest02.png" alt="WriteTest02" style="zoom:80%;" />

### 编写测试方法

写入的时候使用样式还是比较繁琐，实际开发能不使用尽量不要用，或者也可以单独封装成一个方法。注意，构造`Workbook`时不要使用`WorkbookFactory.create(file)`方式，否则，模板也会被修改。

```java
	@Test
	public void test02() throws SQLException, IOException {
		// 处理XSSF
		String templatePath = "extend\\file\\user_data_template.xlsx";
		String outpath = "extend\\file\\user_data.xlsx";

		// 处理HSSF
		// String templatePath = "extend\\file\\user_data_template.xls";
		// String path = "extend\\file\\user_data.xls";

		InputStream in = new FileInputStream(templatePath);

		// 创建工作簿，注意，这里如果传入File对象，模板也会被改写
		Workbook wb = WorkbookFactory.create(in);

		// 读取工作表
		Sheet sheet = wb.getSheetAt(0);

		// 定义复用变量
		int rowIndex = 0; // 行的索引
		int cellIndex = 1; // 单元格的索引
		Row nRow = null;
		Cell nCell = null;

		// 读取大标题行
		nRow = sheet.getRow(rowIndex++); // 使用后 +1
		// 读取大标题的单元格
		nCell = nRow.getCell(cellIndex);
		// 设置大标题的内容
		nCell.setCellValue("2020年2月用户表");

		// 跳过第二行(模板的小标题)
		rowIndex++;

		// 读取第三行,获取它的样式
		nRow = sheet.getRow(rowIndex);
		// 读取行高
		float lineHeight = nRow.getHeightInPoints();
		// 获取第三行的4个单元格中的样式
		CellStyle cs1 = nRow.getCell(cellIndex++).getCellStyle();
		CellStyle cs2 = nRow.getCell(cellIndex++).getCellStyle();
		CellStyle cs3 = nRow.getCell(cellIndex++).getCellStyle();
		CellStyle cs4 = nRow.getCell(cellIndex++).getCellStyle();

		// 查询用户列表
		List<UserDTO> userList = new UserService().findAll().stream().map((x) -> new UserDTO(x)).collect(Collectors.toList());
		// 遍历数据
		for(UserDTO user : userList) {
			// 创建数据行
			nRow = sheet.createRow(rowIndex++);
			// 设置数据行高
			nRow.setHeightInPoints(lineHeight);
			// 重置cellIndex,从第一列开始写数据
			cellIndex = 1;

			// 创建数据单元格，设置单元格内容和样式
			// 用户名
			nCell = nRow.createCell(cellIndex++);
			nCell.setCellStyle(cs1);
			nCell.setCellValue(user.getName());
			// 性别
			nCell = nRow.createCell(cellIndex++);
			nCell.setCellStyle(cs2);
			nCell.setCellValue(user.getGenderStr());
			// 年龄
			nCell = nRow.createCell(cellIndex++);
			nCell.setCellStyle(cs3);
			nCell.setCellValue(user.getAge());
			// 手机号
			nCell = nRow.createCell(cellIndex++);
			nCell.setCellStyle(cs4);
			nCell.setCellValue(user.getPhone());
		}

		// 保存到本地目录
		OutputStream out = new FileOutputStream(new File(outpath));
		wb.write(out);

		// 释放资源
		out.close();
		wb.close();
	}
```

### 测试

运行以上方法，在指定文件夹可以看到生成的文件：

<img src="D:\growUp\git_repository\09-poi-demo\extend\img\WriteTest03.png" alt="WriteTest03" style="zoom:67%;" />

## 读xls--使用SAX方式

### 需求

使用 SAX 的方式将 xls 中的用户数据导入到数据库，数据与以上例子一样。

### 编写测试方法

相比前面的例子，使用 SAX 方式内存占用小，效率高，但是 POI 提供的这套 API 用起来非常繁琐，使用时不得不必须去了解 xls 文件的结构。我这里只是简单展示，监听器部分的代码不太严谨，实际项目还是用 easyexcel 来操作吧。

```java
	@Test
	public void test02() throws Exception {
		// 创建POIFSFileSystem
		String filename = "extend\\file\\user_data.xls";
		POIFSFileSystem poifs = new POIFSFileSystem(new File(filename));

		// 创建HSSFRequest，并添加自定义监听器
		HSSFRequest req = new HSSFRequest();
		EventExample listener = new EventExample();
		req.addListenerForAllRecords(listener);

		// 解析和触发事件
		HSSFEventFactory factory = new HSSFEventFactory();
		factory.processWorkbookEvents(req, poifs);
		
		// 保存用户到数据库
		new UserService().save(listener.getList());
		
		poifs.close();
	}

	private static class EventExample implements HSSFListener {

		private SSTRecord sstrec;

		private int lastCellRow = -1;

		private int lastCellColumn = -1;

		private List<UserDTO> list = new ArrayList<UserDTO>();

		private UserDTO user;

		@Override
		public void processRecord(Record record) {
			switch(record.getSid()) {

			// 进入新的sheet
			case BoundSheetRecord.sid:
				lastCellRow = -1;
				lastCellColumn = -1;
				break;
			
			// excel中的数值类型和字符存放在不同的位置
			case NumberRecord.sid:
				NumberRecord numrec = (NumberRecord)record;
                // 用户年龄
				user.setAge(Double.valueOf(numrec.getValue()).intValue());
                lastCellRow = numrec.getRow();
				lastCellColumn = numrec.getColumn();
				break;

			// SSTRecords中存储着excel中使用的字符，重复的会合并为一个
			case SSTRecord.sid:
				sstrec = (SSTRecord)record;
				break;

			// 读取到单元格的字符
			case LabelSSTRecord.sid:
				LabelSSTRecord lrec = (LabelSSTRecord)record;
				int thisRow = lrec.getRow();
				// 用户数据从第三行开始
				if(thisRow >= 2) {
					// 进入新行时，原对象放入集合，并创建新对象
					if(thisRow != lastCellRow) {
						if(user != null) {
							list.add(user);
						}
						user = new UserDTO();
					}
					// 根据列数为用户对象设置属性
					switch(lrec.getColumn()) {
					case 1:
                        // 用户名
						user.setName(sstrec.getString(lrec.getSSTIndex()).getString());
						break;
					case 2:
                        // 用户性别
						user.setGenderStr(sstrec.getString(lrec.getSSTIndex()).getString());
						break;
					case 4:
                        // 用户电话
						user.setPhone(sstrec.getString(lrec.getSSTIndex()).getString());
						break;
					default:
						break;
					}
					lastCellRow = thisRow;
					lastCellColumn = lrec.getColumn();
				}
				break;
			case EOFRecord.sid:
				// 最后一行读取完后直接放入集合
				if(lastCellRow != -1 && user != null && lastCellColumn == 4) {
					list.add(user);
				}
				break;
			default:
				break;
			}
		}

		public List<UserDTO> getList() {
			return list;
		}
	}
```

### 测试

运行以上方法，可以在数据库看到导入的数据：

<img src="D:\growUp\git_repository\09-poi-demo\extend\img\ReadTest03.png" alt="ReadTest03" style="zoom:80%;" />

## 读xlsx--使用SAX方式

### 需求

使用 SAX 的方式将 xlsx 中的用户数据导入到数据库，数据与以上例子一样。

### 编写测试方法

POI 针对 xlsx 的 SAX API 也是非常繁琐，属于非常低级的封装，这里竟然需要使用 JDK 原生的 SAX 解析来处理事件，定义事件处理器时，我必须去了解 xml 的节点结构。和上面例子一样，这里也只是简单地演示这套 API 的使用，具体代码不太严谨，当然，实际开发我们不会采用这种方式，建议还是使用 easyexcel 吧。

```java
	@Test
	public void test01() throws Exception {

		String filename = "extend\\file\\user_data.xlsx";
		OPCPackage pkg = OPCPackage.open(filename);
		XSSFReader r = new XSSFReader(pkg);

		// 获取sharedStrings.xml的内容，这里存放着excel中的字符
		SharedStringsTable sst = r.getSharedStringsTable();

		// 接下来就是采用SAX方式解析xml的过程
		// 构造解析器，这里会设置自定义的处理器
		XMLReader parser = XMLHelper.newXMLReader();
		SheetHandler handler = new SheetHandler(sst);
		parser.setContentHandler(handler);

		// 解析指定的sheet
		InputStream sheet2 = r.getSheet("rId1");
		parser.parse(new InputSource(sheet2));

		// 保存用户到数据库
		new UserService().save(handler.getList());
		// handler.getList().forEach(System.err::println);

		sheet2.close();
	}

	private static class SheetHandler extends DefaultHandler {

		private SharedStringsTable sst;

		private String cellContents;

		private boolean cellContentsIsString;

		private int cellColumn = -1;

		private int cellRow = -1;

		List<UserDTO> list = new ArrayList<>();

		UserDTO user;

		private SheetHandler(SharedStringsTable sst) {
			this.sst = sst;
		}

		@Override
		public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
			// 读取到行
			if("row".equals(name)) {
				cellRow++;
				if(cellRow >= 2) {
					// 换行时重新创建用户实例
					user = new UserDTO();
				}
			}
			// 读取到列 c => cell
			if("c".equals(name) && cellRow >= 2) {
				// 设置当前读取到哪一列
				char columnChar = attributes.getValue("r").charAt(0);
				switch(columnChar) {
				case 'B':
					cellColumn = 1;
					break;
				case 'C':
					cellColumn = 2;
					break;
				case 'D':
					cellColumn = 3;
					break;
				case 'E':
					cellColumn = 4;
					break;
				default:
					cellColumn = -1;
					break;
				}
				// 当前单元格中的值是否为字符，是的话对应的值被放在SharedStringsTable中
				if("s".equals(attributes.getValue("t"))) {
					cellContentsIsString = true;
				}
			}
			// Clear contents cache
			cellContents = "";
		}

		@Override
		public void endElement(String uri, String localName, String name) throws SAXException {
			// 跳过标题
			if(cellRow < 2) {
				return;
			}
			// v节点是c的子节点，表示单元格的值
			if(name.equals("v")) {
				int idx;
				if(cellContentsIsString) {
					idx = Integer.parseInt(cellContents);
				} else {
					idx = Double.valueOf(cellContents).intValue();
				}
				switch(cellColumn) {
				case 1:
					user.setName(sst.getItemAt(idx).getString());
					break;
				case 2:
					user.setGenderStr(sst.getItemAt(idx).getString());
					break;
				case 3:
					// 年龄的值是数值类型，不在SharedStringsTable中
					user.setAge(idx);
					break;
				case 4:
					user.setPhone(sst.getItemAt(idx).getString());
					break;
				default:
					break;
				}
			}
			
			// 读取完一行，将用户对象放入集合中
			if("row".equals(name) && user != null) {
				list.add(user);
			}
			
			// 重置参数
			if("c".equals(name)) {
				cellColumn = -1;
				cellContentsIsString = false;
			}

		}

		@Override
		public void characters(char[] ch, int start, int length) {
			cellContents += new String(ch, start, length);
		}

		public List<UserDTO> getList() {
			return list;
		}
	}
```

### 测试

运行以上方法，可以在数据库看到导入的数据：

<img src="D:\growUp\git_repository\09-poi-demo\extend\img\ReadTest03.png" alt="ReadTest03" style="zoom:80%;" />

# 源码分析


# 参考资料

[Apache POI - the Java API for Microsoft Documents]( https://poi.apache.org)

>本文为原创文章，转载请附上原文出处链接：https://github.com/ZhangZiSheng001/01-spi-demo
