package cn.zzs.poi;

import java.io.File;
import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

/**
 * <p>测试使用poi读取excel--usermodel</p>
 * @author: zzs
 * @date: 2020年2月25日 下午12:50:38
 */
public class ReadTest {

	/**
	 * <p>入门案例</p>
	 * @throws IOException 
	 */
	@Test
	public void test01() throws IOException {
		// 处理XSSF
		String path = "extend\\file\\poi_test_01.xlsx";
		// 处理HSSF
		// String path = "extend\\file\\poi_test_01.xls";

		// 创建工作簿，会根据excel命名选择不同的Workbook实现类
		Workbook wb = WorkbookFactory.create(new File(path));

		// 获取工作表
		Sheet sheet = wb.getSheetAt(0);

		// 获取行
		Row row = sheet.getRow(0);

		// 获取单元格
		Cell cell = row.getCell(0);
		
		// 也可以采用以下方式获取单元格
		// Cell cell = SheetUtil.getCell(sheet, 0, 0);

		// 获取单元格内容
		String value = cell.getStringCellValue();
		System.err.println("第一个单元格字符：" + value);

		// 释放资源
		wb.close();
	}

	/**
	 * <p>批量导入excel的数据到数据库</p>
	 * @throws SQLException 
	 * @throws IOException 
	 */
	@Test
	public void test02() throws SQLException, IOException {
		// 处理XSSF
		String path = "extend\\file\\user_data.xlsx";
		// 处理HSSF
		// String path = "extend\\file\\user_data.xls";

		// 定义集合，用于存放excel中的用户数据
		List<UserDTO> list = new ArrayList<>();

		// 创建工作簿
		Workbook wb = WorkbookFactory.create(new File(path));

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
		// new UserService().save(list);
		list.forEach(System.err::println);

		// 释放资源
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
		// // 用户名
		// cell = row.getCell(1);
		// userDTO.setName((String)getValue(cell));
		// // 性别
		// cell = row.getCell(2);
		// userDTO.setGenderStr((String)getValue(cell));
		// // 年龄
		// cell = row.getCell(3);
		// userDTO.setAge(((Double)getValue(cell)).intValue());
		// // 电话
		// cell = row.getCell(4);
		// userDTO.setPhone((String)getValue(cell));

		return userDTO;
	}

	/**
	 * <p>获取单元格内的数据,并进行格式转换</p>
	 */
	@SuppressWarnings("unused")
	private Object getValue(Cell cell) {
		switch(cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue();
		case BOOLEAN:
			return cell.getBooleanCellValue();
		case NUMERIC:// 数值和日期均是此类型
			if(DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			} else {
				return cell.getNumericCellValue();
			}
		case FORMULA:
			return cell.getCellFormula();
		default:
			return null;
		}
	}
}
