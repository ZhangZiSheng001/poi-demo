package cn.zzs.easyexcel;

import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.listener.ReadListener;

import cn.zzs.poi.UserDTO;
import cn.zzs.poi.UserService;

/**
 * <p>测试使用easyexcel读取excel</p>
 * @author: zzs
 * @date: 2020年2月25日 下午12:50:38
 */
public class ReadTest {

	/**
	 * <p>批量导入excel的数据到数据库--使用自定义监听器</p>
	 * @throws SQLException 
	 * @throws IOException 
	 */
	@Test
	public void test01() throws SQLException, IOException {
		// XSSF
		String path = "D:\\growUp\\git_repository\\09-poi-demo\\extend\\file\\user_data.xlsx";
		
		// HSSF
		// String path = "D:\\growUp\\git_repository\\09-poi-demo\\extend\\file\\user_data.xls";

		List<UserDTO> list = new ArrayList<UserDTO>();
		
		// 定义回调监听器
		ReadListener<UserDTO> syncReadListener = new AnalysisEventListener<UserDTO>() {
			@Override
			public void invoke(UserDTO data, AnalysisContext context) {
				list.add(data);
			}

			@Override
			public void doAfterAllAnalysed(AnalysisContext context) {
				// TODO Auto-generated method stub
				
			}
		};
		// 读取excel
		EasyExcel.read(path, UserDTO.class, syncReadListener).sheet(0).headRowNumber(2).doRead();
		// 保存
		new UserService().save(list);
	}

	/**
	 * <p>批量导入excel的数据到数据库--使用默认监听器</p>
	 * @throws SQLException 
	 * @throws IOException 
	 */
	@Test
	public void test02() throws SQLException, IOException {
		// XSSF
		String path = "D:\\growUp\\git_repository\\09-poi-demo\\extend\\file\\user_data.xlsx";
		// HSSF
		// String path = "D:\\growUp\\git_repository\\09-poi-demo\\extend\\file\\user_data.xls";

		// 读取excel
		List<UserDTO> list = EasyExcel.read(path).head(UserDTO.class).sheet(0).headRowNumber(2).doReadSync();
		// 保存
		new UserService().save(list);
	}
}