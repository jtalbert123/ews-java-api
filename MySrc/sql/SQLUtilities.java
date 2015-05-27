package sql;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

public class SQLUtilities {

	private static final boolean testing = false;

	static Connection connection = getNewConnection();
	public static Statement statement = getNewStatement();

	private static Connection getNewConnection() {
		Connection connect;
		try {
			Class.forName("com.mysql.jdbc.Driver");
			if (testing)
				connect = DriverManager
						.getConnection("jdbc:mysql://passwordchanger/test?user=dbUser&password=dbusr");
			else
				connect = DriverManager
						.getConnection("jdbc:mysql://passwordchanger/production?user=dbUser&password=dbusr");

			return connect;

		} catch (SQLException e) {
			e.printStackTrace();
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		}
		return null;
	}

	public static Statement getNewStatement() {
		try {
			if (connection == null)
				connection = getNewConnection();
			Statement s = connection.createStatement();

			return s;
		} catch (SQLException e) {

		}
		return null;
	}

	public static PreparedStatement getNewPreparedStatement(String format) {
		try {
			if (connection == null)
				connection = getNewConnection();
			PreparedStatement s = connection.prepareStatement(format);

			return s;
		} catch (SQLException e) {

		}
		return null;
	}

	public static Statement getStatement() {
		return statement;
	}

	/**
	 * Clears all rows from the specified table, if it exists.
	 * 
	 * @return if the table was cleared.
	 */
	public static boolean clearTable(String tableName) {
		try {
			statement.execute("delete from " + tableName + ";");
			statement.execute("ALTER TABLE " + tableName
					+ " AUTO_INCREMENT = 0;");
			return true;
		} catch (SQLException e) {
			e.printStackTrace();
			return false;
		}
	}

	public static void printSet(ResultSet set, int columns) throws SQLException {
		while (set.next()) {
			for (int i = 1; i <= columns; i++) {
				System.out.print(set.getString(i) + ", ");
			}
			System.out.println();
		}
	}

}
