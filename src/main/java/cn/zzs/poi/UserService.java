package cn.zzs.poi;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

/**
 * <p>用户 服务类（简化）</p>
 * @author: zzs
 * @date: 2020年2月24日 下午3:31:06
 */
public class UserService {

	/**
	 * 
	 * <p>保存用户</p>
	 * @author: zzs
	 * @date: 2020年2月25日 上午10:48:31
	 * @param user
	 * @return: void
	 * @throws Exception 
	 */
	public void save(UserDTO user) throws SQLException {
		// 创建sql
		String sql = "insert into demo_user(`id`,`name`,`gender`,`age`,`gmt_create`,`gmt_modified`,`deleted`,`phone`) values(REPLACE(UUID(),'-',''),?,?,?,now(),now(),?,?)";
		Connection connection = null;
		PreparedStatement statement = null;
		try {
			// 获得连接
			connection = JDBCUtils.getConnection();
			// 开启事务设置非自动提交
			JDBCUtils.startTrasaction();
			// 获得Statement对象
			statement = connection.prepareStatement(sql);
			// 设置参数
			setParamToStatement(user, statement);
			// 执行
			statement.executeUpdate();
			// 提交事务
			JDBCUtils.commit();
		} catch(Exception e) {
			JDBCUtils.rollback();
			throw e;
		} finally {
			// 释放资源
			JDBCUtils.release(connection, statement, null);
		}
	}

	/**
	 * 
	 * <p>批量保存用户</p>
	 * @author: zzs
	 * @date: 2020年2月25日 上午10:49:12
	 * @param list
	 * @return: void
	 * @throws Exception 
	 */
	public void save(List<UserDTO> list) throws SQLException {
		// 创建sql
		String sql = "insert into demo_user(`id`,`name`,`gender`,`age`,`gmt_create`,`gmt_modified`,`deleted`,`phone`) values (REPLACE(UUID(),'-',''),?,?,?,now(),now(),?,?)";
		Connection connection = null;
		PreparedStatement statement = null;
		try {
			// 获得连接
			connection = JDBCUtils.getConnection();
			// 开启事务设置非自动提交
			JDBCUtils.startTrasaction();
			// 获得Statement对象
			statement = connection.prepareStatement(sql);
			// 设置参数
			for(int i = 0; i < list.size(); i++) {
				UserDTO user = list.get(i);
				setParamToStatement(user, statement);
				statement.addBatch();
			}
			// 执行
			statement.executeBatch();
			// 提交事务
			JDBCUtils.commit();
		} catch(Exception e) {
			JDBCUtils.rollback();
			throw e;
		} finally {
			// 释放资源
			JDBCUtils.release(connection, statement, null);
		}
	}

	/**
	 * 
	 * <p>查询所有用户（未删除）</p>
	 * @author: zzs
	 * @date: 2020年2月25日 上午11:15:50
	 * @return: List<User>
	 * @throws SQLException 
	 */
	public List<User> findAll() throws SQLException {
		// 创建sql
		String sql = "select * from demo_user where deleted = false";
		Connection connection = null;
		PreparedStatement statement = null;
		ResultSet resultSet = null;
		List<User> list = new ArrayList<>();
		try {
			// 获得连接
			connection = JDBCUtils.getConnection();
			// 获得Statement对象
			statement = connection.prepareStatement(sql);
			// 执行
			resultSet = statement.executeQuery();
			// 遍历结果集
			while(resultSet.next()) {
				list.add(convertResultToUser(resultSet));
			}
			return list;
		} finally {
			// 释放资源
			JDBCUtils.release(connection, statement, resultSet);
		}
	}

	/**
	 * 
	 * <p>根据结果集封装用户对象</p>
	 * @author: zzs
	 * @date: 2020年2月25日 上午11:22:13
	 * @param resultSet
	 * @return: User
	 * @throws SQLException 
	 */
	private User convertResultToUser(ResultSet resultSet) throws SQLException {
		User user = new User();
		user.setId(resultSet.getString(1));
		user.setName(resultSet.getString(2));
		user.setGender(resultSet.getInt(3));
		user.setAge(resultSet.getInt(4));
		user.setGmtCreate(resultSet.getDate(5));
		user.setGmtModified(resultSet.getDate(6));
		user.setDeleted(resultSet.getInt(7));
		user.setPhone(resultSet.getString(8));
		return user;
	}

	/**
	 * 
	 * <p>设置语句的参数</p>
	 * @author: zzs
	 * @date: 2020年2月25日 上午11:00:13
	 * @param user
	 * @param statement
	 * @throws SQLException
	 * @return: void
	 */
	private void setParamToStatement(UserDTO user, PreparedStatement statement) throws SQLException {
		statement.setString(1, user.getName());
		statement.setInt(2, user.getGender());
		statement.setInt(3, user.getAge());
		statement.setInt(4, 0);
		statement.setString(5, user.getPhone());
	}
}
