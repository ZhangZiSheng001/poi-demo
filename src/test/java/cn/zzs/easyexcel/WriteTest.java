package cn.zzs.easyexcel;

import java.io.IOException;
import java.sql.SQLException;
import java.util.List;
import java.util.stream.Collectors;

import org.junit.Test;

import com.alibaba.excel.EasyExcel;

import cn.zzs.poi.UserDTO;
import cn.zzs.poi.UserService;

/**
 * <p>测试使用easyexcel写入excel</p>
 * @author: zzs
 * @date: 2020年2月23日 上午9:20:03
 */
public class WriteTest {

    /**
     * <p>批量写入用户数据到excel</p>
     * @throws SQLException 
     * @throws IOException 
     */
    @Test
    public void test01() throws SQLException, IOException {
        // XSSF
        String path = "D:\\growUp\\git_repository\\09-poi-demo\\extend\\file\\user_data.xlsx";

        // HSSF
        // String path = "D:\\growUp\\git_repository\\09-poi-demo\\extend\\file\\user_data.xls";

        // 获取用户数据
        List<UserDTO> list = new UserService().findAll().stream().map((x) -> new UserDTO(x)).collect(Collectors.toList());
        // 写入excel
        EasyExcel.write(path, UserDTO.class).sheet(0).relativeHeadRowIndex(1).doWrite(list);
    }

}
