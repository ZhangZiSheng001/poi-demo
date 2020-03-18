package cn.zzs.poi;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.EOFRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.junit.Test;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

/**
 * <p>测试使用poi读取excel--eventmodel</p>
 * @author: zzs
 * @date: 2020年3月17日 上午11:17:22
 */
public class SAXReadTest {

	/**
	 * 
	 * <p>使用SAX方式读取xlsx</p>
	 */
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
		// new UserService().save(handler.getList());
		handler.getList().forEach(System.err::println);

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

	/**
	 * 
	 * <p>使用SAX方式读取xls</p>
	 */
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
		//new UserService().save(listener.getList());
		listener.getList().forEach(System.err::println);

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
						user.setName(sstrec.getString(lrec.getSSTIndex()).getString());
						break;
					case 2:
						user.setGenderStr(sstrec.getString(lrec.getSSTIndex()).getString());
						break;
					case 4:
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
}
