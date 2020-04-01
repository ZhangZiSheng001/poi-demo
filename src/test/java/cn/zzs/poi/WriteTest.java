package cn.zzs.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.SQLException;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.WorkbookUtil;
import org.junit.Test;

/**
 * <p>测试使用poi写入excel</p>
 * @author: zzs
 * @date: 2020年2月23日 上午9:20:03
 */
public class WriteTest {

    /**
     * <p>入门案例</p>
     */
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

    /**
     * <p>使用模板批量写入用户数据到excel</p>
     * @throws SQLException 
     * @throws IOException 
     */
    @Test
    public void test02() throws SQLException, IOException {
        // 处理XSSF
        String templatePath = "extend\\file\\user_data_template.xlsx";
        String outpath = "extend\\file\\user_data.xlsx";

        // 处理HSSF
        // String templatePath = "extend\\file\\user_data_template.xls";
        // String outpath = "extend\\file\\user_data.xls";

        InputStream in = new FileInputStream(templatePath);

        // 创建工作簿
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
        in.close();
        out.close();
        wb.close();
    }

}
