package com.excelbean;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.sql.Blob;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

public class ExcelBrowser {

	private final Log log = LogFactory.getLog(ExcelBrowser.class);

	// Define the variables
	// public static final String filePath = "D:/tools/excelmanager";
	public static String uploadTempFilePath = "D:/tools/excelmanager";

	public static final String tableNamePrefix = "sheet";

	public static final String datePattern = "YY-MM-DD hh:mm:ss";

	public static String mysqlUrl = "jdbc:mysql://127.0.0.1:3306/excelmanager";
	public static String mysqlUser = "root";
	public static String mysqlPassword = "root";

	public static final String XLSX = "xlsx";
	public static final String XLS = "xls";

	public String saveExcelFile(String extendedName, byte[] content) {
		Connection connection = null;
		PreparedStatement preparedStatement = null;

		int id = 0;

		try {
			connection = DriverManager.getConnection(mysqlUrl, mysqlUser, mysqlPassword);

			StringBuilder insertSql = new StringBuilder();
			insertSql.append("insert into excelfile (extendname, data) values (?, ?)");

			preparedStatement = connection.prepareStatement(insertSql.toString());

			preparedStatement.setString(1, extendedName);
			preparedStatement.setBinaryStream(2, new ByteArrayInputStream(content));

			preparedStatement.execute();

			ResultSet resultSet = preparedStatement.executeQuery("select last_insert_id() as id");
			while (resultSet.next()) {
				id = resultSet.getInt(1);

				log.warn("Excel file id:" + id);

				break;
			}

			resultSet.close();

		} catch (Exception e) {
			log.warn(e);
		} finally {
			try {
				if (preparedStatement != null) {
					preparedStatement.close();
				}

				if (connection != null) {
					connection.close();
				}
			} catch (Exception e) {
				log.warn(e);
			}
		}

		String sheetsJson;

		if (extendedName.compareTo(XLSX) == 0) {
			sheetsJson = xlsxGetSheetsJson(id, content);
		} else {
			sheetsJson = xlsGetSheetsJson(id, content);

			// JSONObject jsonObject = new JSONObject();

			// jsonObject.put("excelFileId", id);

			// JSONArray jsonArray = new JSONArray();

			// jsonObject.put("sheetNames", jsonArray);

			// sheetsJson = jsonObject.toJSONString();
		}

		return sheetsJson;
	}

	public void exelToMySql(int excelFileId, List<Map<String, String>> sheetNameNotes) {
		Connection connection = null;
		PreparedStatement preparedStatement = null;

		InputStream inputStream = null;

		try {
			connection = DriverManager.getConnection(mysqlUrl, mysqlUser, mysqlPassword);

			StringBuilder insertSql = new StringBuilder();
			insertSql.append("select extendname,data from excelfile where id = ?");

			preparedStatement = connection.prepareStatement(insertSql.toString());

			preparedStatement.setInt(1, excelFileId);

			ResultSet resultSet = preparedStatement.executeQuery();
			while (resultSet.next()) {
				String extendName = resultSet.getString(1);

				Blob blob = resultSet.getBlob(2);

				inputStream = blob.getBinaryStream();

				if (extendName.compareToIgnoreCase(XLSX) == 0) {

					xlsxToMySql(inputStream, sheetNameNotes);

				} else {

					xlsToMySql(inputStream, sheetNameNotes);

				}

				break;
			}

			resultSet.close();

		} catch (Exception e) {
			log.warn(e);
		} finally {
			try {
				if (preparedStatement != null) {
					preparedStatement.close();
				}

				if (connection != null) {
					connection.close();
				}
			} catch (Exception e) {
				log.warn(e);
			}
		}

	}

	public String listAllSheets() {
		JSONObject jsonObject = new JSONObject();

		JSONArray jsonArray = new JSONArray();

		Connection connection = null;
		Statement statement = null;

		try {
			connection = DriverManager.getConnection(mysqlUrl, mysqlUser, mysqlPassword);

			String querySql = "select id, tablename, sheetname, alias, description, department, records, timestamp from sheetinformation";

			statement = connection.createStatement();

			ResultSet resultSet = statement.executeQuery(querySql);

			while (resultSet.next()) {
				JSONObject sheetObject = resultSetToSheet(resultSet);

				jsonArray.add(sheetObject);
			}

			resultSet.close();

		} catch (Exception e) {
			log.warn(e);
		} finally {
			try {
				if (statement != null) {
					statement.close();
				}

				if (connection != null) {
					connection.close();
				}
			} catch (Exception e) {
				log.warn(e);
			}
		}

		jsonObject.put("sheets", jsonArray);

		return jsonObject.toJSONString();
	}

	public void delSheets(List<Integer> sheetIds) {
		Connection connection = null;
		Statement statement = null;

		try {
			connection = DriverManager.getConnection(mysqlUrl, mysqlUser, mysqlPassword);
			statement = connection.createStatement();

			for (int sheetId : sheetIds) {
				String querySql = "select tablename from sheetinformation where id = " + sheetId;

				ResultSet resultSet = statement.executeQuery(querySql);

				String tableName = "";

				while (resultSet.next()) {

					tableName = resultSet.getString(1);

					break;
				}

				resultSet.close();

				String deleteRecordsSql = "delete from " + tableName;
				statement.execute(deleteRecordsSql);

				String dropTableSql = "drop table if exists " + tableName;
				statement.execute(dropTableSql);

				String deleteColumnInformationSql = "delete from columninformation where sheetinformationid = "
						+ sheetId;
				statement.execute(deleteColumnInformationSql);

				String deleteSheetInformationSql = "delete from sheetinformation where id = " + sheetId;
				statement.execute(deleteSheetInformationSql);
			}

		} catch (Exception e) {
			log.warn(e);
		} finally {
			try {
				if (statement != null) {
					statement.close();
				}

				if (connection != null) {
					connection.close();
				}
			} catch (Exception e) {
				log.warn(e);
			}
		}

	}

	public String previewSheets(List<Integer> sheetIds) {

		JSONObject jsonObject = sheetsToJsonObject(sheetIds);

		return jsonObject.toJSONString();

	}

	public XSSFWorkbook downloadSheets(List<Integer> sheetIds) {

		JSONObject jsonObject = sheetsToJsonObject(sheetIds);

		XSSFWorkbook xssfWorkbook = sheetsJsonObjectToXlsxFile(jsonObject);

		return xssfWorkbook;
	}

	public void editSheet(int sheetId, String alias, String description) {
		Connection connection = null;
		PreparedStatement preparedStatement = null;

		JSONObject jsonObject = new JSONObject();

		JSONArray sheetJsonArray = new JSONArray();

		try {
			connection = DriverManager.getConnection(mysqlUrl, mysqlUser, mysqlPassword);

			String updateSql = "update sheetinformation set alias = ?, description = ? where id = ?";

			preparedStatement = connection.prepareStatement(updateSql);

			preparedStatement.setString(1, alias);
			preparedStatement.setString(2, description);
			preparedStatement.setInt(3, sheetId);

			preparedStatement.executeUpdate();

		} catch (Exception e) {
			log.warn(e);
		} finally {
			try {
				if (preparedStatement != null) {
					preparedStatement.close();
				}

				if (connection != null) {
					connection.close();
				}
			} catch (Exception e) {
				log.warn(e);
			}
		}
	}

	public String searchSheets(String alias, String description) {
		JSONObject jsonObject = new JSONObject();

		JSONArray jsonArray = new JSONArray();

		Connection connection = null;
		Statement statement = null;

		StringBuilder condition = new StringBuilder();

		if (!StringUtils.isEmpty(alias) && StringUtils.isEmpty(description)) {

			condition.append(" where alias like '%");
			condition.append(alias);
			condition.append("%'");

		} else if (StringUtils.isEmpty(alias) && !StringUtils.isEmpty(description)) {

			condition.append(" where description like '%");
			condition.append(description);
			condition.append("%'");

		} else if (!StringUtils.isEmpty(alias) && !StringUtils.isEmpty(description)) {

			condition.append(" where alias like '%");
			condition.append(alias);
			condition.append("%'");

			condition.append(" or");
			condition.append(" description like '%");
			condition.append(description);
			condition.append("%'");

		}

		try {
			connection = DriverManager.getConnection(mysqlUrl, mysqlUser, mysqlPassword);

			String querySql = "select id, tablename, sheetname, alias, description, department, records, timestamp from sheetinformation"
					+ condition.toString();

			statement = connection.createStatement();

			ResultSet resultSet = statement.executeQuery(querySql);

			while (resultSet.next()) {
				JSONObject sheetObject = resultSetToSheet(resultSet);

				jsonArray.add(sheetObject);
			}

			resultSet.close();

		} catch (Exception e) {
			log.warn(e);
		} finally {
			try {
				if (statement != null) {
					statement.close();
				}

				if (connection != null) {
					connection.close();
				}
			} catch (Exception e) {
				log.warn(e);
			}
		}

		jsonObject.put("sheets", jsonArray);

		return jsonObject.toJSONString();

	}

	public String sheetCount() {
		JSONObject jsonObject = new JSONObject();

		JSONArray jsonArray = new JSONArray();

		Connection connection = null;
		Statement statement = null;

		try {
			connection = DriverManager.getConnection(mysqlUrl, mysqlUser, mysqlPassword);

			String querySql = "select count(*) from sheetinformation";

			statement = connection.createStatement();

			ResultSet resultSet = statement.executeQuery(querySql);

			while (resultSet.next()) {
				int sheetCount = resultSet.getInt(1);

				jsonObject.put("sheetCount", sheetCount);

				break;
			}

			resultSet.close();

		} catch (Exception e) {
			log.warn(e);
		} finally {
			try {
				if (statement != null) {
					statement.close();
				}

				if (connection != null) {
					connection.close();
				}
			} catch (Exception e) {
				log.warn(e);
			}
		}

		return jsonObject.toJSONString();

	}

	public String pagingSheets(int pageNumber, int sheetPerPage) {
		JSONObject jsonObject = new JSONObject();

		JSONArray jsonArray = new JSONArray();

		Connection connection = null;
		Statement statement = null;

		StringBuilder limit = new StringBuilder();

		limit.append(" limit ");
		limit.append(pageNumber);
		limit.append(",");
		limit.append(sheetPerPage);

		try {
			connection = DriverManager.getConnection(mysqlUrl, mysqlUser, mysqlPassword);

			statement = connection.createStatement();

			String querySql = "select id, tablename, sheetname, alias, description, department, records, timestamp from sheetinformation"
					+ limit.toString();

			log.warn(querySql);

			ResultSet resultSet = statement.executeQuery(querySql);

			while (resultSet.next()) {
				JSONObject sheetObject = resultSetToSheet(resultSet);

				jsonArray.add(sheetObject);
			}

			resultSet.close();

		} catch (Exception e) {
			log.warn(e);
		} finally {
			try {
				if (statement != null) {
					statement.close();
				}

				if (connection != null) {
					connection.close();
				}
			} catch (Exception e) {
				log.warn(e);
			}
		}

		jsonObject.put("sheets", jsonArray);

		return jsonObject.toJSONString();

	}

	private JSONObject sheetsToJsonObject(List<Integer> sheetIds) {
		Connection connection = null;
		Statement statement = null;

		JSONObject jsonObject = new JSONObject();

		JSONArray sheetJsonArray = new JSONArray();

		try {
			connection = DriverManager.getConnection(mysqlUrl, mysqlUser, mysqlPassword);
			statement = connection.createStatement();

			for (int sheetId : sheetIds) {

				String querySql = "select id, tablename, sheetname, alias, description, department, records, timestamp from sheetinformation where id = "
						+ sheetId;

				ResultSet resultSet = statement.executeQuery(querySql);

				JSONObject sheetJsonObject = new JSONObject();

				String tableName = "";
				while (resultSet.next()) {

					JSONObject sheetInformationJsonObject = resultSetToSheet(resultSet);

					tableName = sheetInformationJsonObject.getString("tableName");

					sheetJsonObject.put("sheetinformation", sheetInformationJsonObject);

					break;
				}

				resultSet.close();

				JSONArray columnInformationJsonArray = getColumnInformation(sheetId, connection, statement);
				sheetJsonObject.put("columninformation", columnInformationJsonArray);

				JSONArray recordsJsonArray = getRecords(columnInformationJsonArray.size(), tableName, connection,
						statement);
				sheetJsonObject.put("records", recordsJsonArray);

				sheetJsonArray.add(sheetJsonObject);

			}

		} catch (Exception e) {
			log.warn(e);
		} finally {
			try {
				if (statement != null) {
					statement.close();
				}

				if (connection != null) {
					connection.close();
				}
			} catch (Exception e) {
				log.warn(e);
			}
		}

		jsonObject.put("sheets", sheetJsonArray);

		return jsonObject;
	}

	private XSSFWorkbook sheetsJsonObjectToXlsxFile(JSONObject sheetsJsonObject) {

		XSSFWorkbook xssfWorkbook = null;

		try {

			// Open the input stream from the file, initialize the workbook

			xssfWorkbook = new XSSFWorkbook();

			JSONArray sheetsJsonArray = sheetsJsonObject.getJSONArray("sheets");

			for (int i = 0; i < sheetsJsonArray.size(); i++) {

				JSONObject sheetJsonObject = sheetsJsonArray.getJSONObject(i);

				JSONObject sheetinformationJsonObject = sheetJsonObject.getJSONObject("sheetinformation");

				String sheetName = sheetinformationJsonObject.getString("alias");

				XSSFSheet xssfSheet = xssfWorkbook.createSheet(sheetName);

				XSSFRow headerRow = xssfSheet.createRow(0);

				JSONArray columinformatonJsonArray = sheetJsonObject.getJSONArray("columninformation");
				for (int j = 0; j < columinformatonJsonArray.size(); j++) {

					JSONObject columnJsonObject = columinformatonJsonArray.getJSONObject(j);

					String nameInExcel = columnJsonObject.getString("nameInExcel");

					XSSFCell xssfCell = headerRow.createCell(j);

					xssfCell.setCellValue(nameInExcel);

				}

				JSONArray recordsJsonArray = sheetJsonObject.getJSONArray("records");
				for (int j = 0; j < recordsJsonArray.size(); j++) {

					XSSFRow dataRow = xssfSheet.createRow(j + 1);

					JSONArray recordJsonArray = recordsJsonArray.getJSONArray(j);

					for (int k = 0; k < recordJsonArray.size(); k++) {

						String value = recordJsonArray.getString(k);

						XSSFCell xssfCell = dataRow.createCell(k);

						xssfCell.setCellValue(value);

					}

				}

			}

		} catch (Exception e) {

			e.printStackTrace();

		} finally {

		}

		return xssfWorkbook;

	}

	private void xlsxToMySql(InputStream inputStream, List<Map<String, String>> sheetNameNotes) {

		try {
			XSSFWorkbook xssfWorkbook = new XSSFWorkbook(inputStream);

			excelToMySql(xssfWorkbook, sheetNameNotes);
		} catch (Exception e) {
			log.warn(e.toString());
		}
	}

	private void xlsToMySql(InputStream inputStream, List<Map<String, String>> sheetNameNotes) {

		try {
			HSSFWorkbook xssfWorkbook = new HSSFWorkbook(inputStream);

			excelToMySql(xssfWorkbook, sheetNameNotes);
		} catch (Exception e) {
			log.warn(e.toString());
		}
	}

	private void excelToMySql(Workbook workbook, List<Map<String, String>> sheetNameNotes) {

		FormulaEvaluator formulaEvaluator;

		int maxCellCount = 0;

		try {

			// Open the input stream from the file, initialize the workbook

			formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

			// Open the sheet

			for (int dataSheetIndex = 0; dataSheetIndex < workbook.getNumberOfSheets(); dataSheetIndex++) {

				Sheet sheet = workbook.getSheetAt(Integer.valueOf(dataSheetIndex));

				if (sheet == null) {
					break;
				}

				// Get the first row of the sheet

				Row firstRow = sheet.getRow(0);

				if (firstRow == null) {
					continue;
				}

				String tableName = createMySqlTable(sheet, firstRow, sheetNameNotes);
				if (tableName == null) {
					continue;
				}

				// Set the max cell count to be the last cell number of
				// the
				// first line

				maxCellCount = firstRow.getLastCellNum();
				if (maxCellCount == 0) {
					continue;
				}

				writeAllRecordToTable(tableName, sheet, formulaEvaluator, maxCellCount);

			}

		} catch (Exception e) {

			e.printStackTrace();

		} finally {

		}

	}

	private String xlsxGetSheetsJson(int excelFileId, byte[] content) {

		XSSFWorkbook xssfWorkbook = null;

		ByteArrayInputStream inputStream = new ByteArrayInputStream(content);

		try {
			xssfWorkbook = new XSSFWorkbook(inputStream);
		} catch (Exception e) {
			log.warn("Input stream can not be loaded as content of excel file.");
		}

		return excelGetSheetsJson(xssfWorkbook, excelFileId, content);

	}

	private String xlsGetSheetsJson(int excelFileId, byte[] content) {

		HSSFWorkbook hssfWorkbook = null;

		ByteArrayInputStream inputStream = new ByteArrayInputStream(content);

		try {
			hssfWorkbook = new HSSFWorkbook(inputStream);
		} catch (Exception e) {
			log.warn("Input stream can not be loaded as content of excel file.");
		}

		return excelGetSheetsJson(hssfWorkbook, excelFileId, content);

	}

	private String excelGetSheetsJson(Workbook workbook, int excelFileId, byte[] content) {

		JSONObject jsonObject = new JSONObject();

		jsonObject.put("excelFileId", excelFileId);

		JSONArray jsonArray = new JSONArray();

		jsonObject.put("sheetNames", jsonArray);

		if (workbook == null) {
			return jsonObject.toJSONString();
		}

		try {

			// Open the sheet

			for (int dataSheetIndex = 0;; dataSheetIndex++) {
				Sheet sheet = null;

				try {
					sheet = workbook.getSheetAt(Integer.valueOf(dataSheetIndex));
				} catch (Exception e) {
					// Datasheet index is outof range, do nothing here.
				}

				if (sheet == null) {
					break;
				}

				jsonArray.add(sheet.getSheetName());
			}

		} catch (Exception e) {

			e.printStackTrace();

		} finally {

		}

		jsonObject.put("sheetNames", jsonArray);

		return jsonObject.toJSONString();
	}

	private String createMySqlTable(Sheet sheet, Row firstRow, List<Map<String, String>> sheetNameNotes) {

		int maxCellCount = firstRow.getLastCellNum();

		List<String> colNames = getColNames(firstRow, maxCellCount);
		if (colNames.size() == 0) {
			return null;
		}

		String sheetName = sheet.getSheetName();
		String tableName = "";
		long timestamp = System.currentTimeMillis();

		StringBuilder insertSql = new StringBuilder();

		insertSql.append(
				"insert into sheetinformation (tablename, sheetname, alias, description, department, records, timestamp) values (");
		insertSql.append("'" + tableName + "', ");
		insertSql.append("'" + sheetName + "', ");

		for (Map<String, String> sheetNameNote : sheetNameNotes) {

			String sheetNameTemp = sheetNameNote.get("sheetName");

			if (sheetNameTemp.equals(sheetName) == true) {

				String alias = sheetNameNote.get("alias");
				insertSql.append("'" + alias + "', ");

				String descriptioin = sheetNameNote.get("description");
				insertSql.append("'" + descriptioin + "', ");

				// For department
				insertSql.append("'', ");

				break;

			}

		}

		insertSql.append("" + sheet.getLastRowNum() + ", ");

		insertSql.append("" + timestamp + ")");

		log.warn(insertSql.toString());

		int id = createWorksheetInformation(insertSql.toString());

		tableName = getTableNameFromId(id);

		Connection connection = null;
		Statement statement = null;

		try {
			connection = DriverManager.getConnection(mysqlUrl, mysqlUser, mysqlPassword);

			statement = connection.createStatement();

			StringBuilder createTableSql = new StringBuilder();
			createTableSql.append("create table `" + tableName + "`(");
			createTableSql.append("`id` int(11) NOT NULL AUTO_INCREMENT,");

			int colIndex = 0;
			for (String colName : colNames) {

				createTableSql.append("`Col" + colIndex + "` VARCHAR(512),");

				StringBuilder stringBuilder = new StringBuilder();

				stringBuilder.append("insert into columninformation (sheetinformationid,colname,nameinexcel) values (");
				stringBuilder.append("" + id + ", ");
				stringBuilder.append("'Col" + colIndex + "', ");
				stringBuilder.append("'" + colName + "')");

				createColumnInformation(stringBuilder.toString(), connection, statement);

				colIndex++;

			}

			createTableSql.append("PRIMARY KEY (`id`)");
			createTableSql.append(") ENGINE=InnoDB DEFAULT CHARSET=utf8");

			log.warn(createTableSql.toString());

			statement.execute(createTableSql.toString());

		} catch (Exception e) {
		} finally {
			try {
				if (statement != null) {
					statement.close();
				}
				if (connection != null) {
					connection.close();
				}
			} catch (Exception e) {

			}
		}

		return tableName;
	}

	private void writeAllRecordToTable(String tableName, Sheet sheet, FormulaEvaluator formulaEvaluator,
			int maxCellCount) {

		Connection connection = null;
		Statement statement = null;
		// Loop to read the rows

		try {
			connection = DriverManager.getConnection(mysqlUrl, mysqlUser, mysqlPassword);

			statement = connection.createStatement();

			for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {

				// Get the row

				Row row = sheet.getRow(rowNum);

				if (row == null) {

					continue;

				}

				insertRowToMySqlTable(tableName, sheet, row, formulaEvaluator, maxCellCount, connection, statement);
			}
		} catch (Exception e) {
		} finally {
			try {
				if (statement != null) {
					statement.close();
				}
				if (connection != null) {
					connection.close();
				}
			} catch (Exception e) {

			}
		}

	}

	private void insertRowToMySqlTable(String tableName, Sheet sheet, Row row, FormulaEvaluator formulaEvaluator,
			int maxCellCount, Connection connection, Statement statement) {

		StringBuilder insertSql = new StringBuilder();
		insertSql.append("insert into " + tableName + " (");

		for (int i = 0; i < maxCellCount; i++) {
			insertSql.append("Col" + i);

			if (i != maxCellCount - 1) {
				insertSql.append(",");
			}
		}

		insertSql.append(") values (");

		// Loop to read the cells

		for (int cellNum = 0; cellNum < maxCellCount; cellNum++) {

			// Get the cell

			Cell xssfCell = row.getCell(cellNum);

			if (xssfCell == null) {
				continue;
			}

			// Process the cell based on the cell
			// type

			switch (xssfCell.getCellTypeEnum()) {

			case STRING:

				insertSql.append("'" + xssfCell.getStringCellValue() + "'");
				// System.out.print(xssfCell.getStringCellValue());

				break;

			case NUMERIC:

				// If the cell matches the date
				// format,
				// output the cell as a date

				if (DateUtil.isCellDateFormatted(xssfCell)) {

					SimpleDateFormat dateFormat = new SimpleDateFormat(datePattern);

					insertSql.append("'" + dateFormat.format(xssfCell.getDateCellValue()) + "'");
					// System.out.print(dateFormat.format(xssfCell.getDateCellValue()));

				}

				else {
					insertSql.append("'" + xssfCell.getNumericCellValue() + "'");
					// System.out.print(xssfCell.getNumericCellValue());
				}

				break;

			case BOOLEAN:

				insertSql.append("'" + xssfCell.getBooleanCellValue() + "'");
				// System.out.print(xssfCell.getBooleanCellValue());

				break;

			case FORMULA:

				// For formula cell, evaluate the
				// formula to get the result

				CellValue cellValue = formulaEvaluator.evaluate(xssfCell);

				// Process the formula cell based on
				// the
				// type of the result

				switch (cellValue.getCellTypeEnum()) {

				case STRING:

					insertSql.append("'" + xssfCell.getStringCellValue() + "'");
					// System.out.print(xssfCell.getStringCellValue());

					break;

				case NUMERIC:

					// If the result matches the
					// date
					// format, output the result as
					// a
					// date

					if (DateUtil.isCellDateFormatted(xssfCell)) {

						SimpleDateFormat dateFormat = new SimpleDateFormat(datePattern);

						insertSql.append("'" + dateFormat.format(xssfCell.getDateCellValue()) + "'");
						// System.out.print(dateFormat.format(xssfCell.getDateCellValue()));

					}

					else {

						insertSql.append("'" + xssfCell.getNumericCellValue() + "'");
						// System.out.print(xssfCell.getNumericCellValue());
					}
					break;

				case BOOLEAN:

					insertSql.append("'" + xssfCell.getBooleanCellValue() + "'");
					// System.out.print(xssfCell.getBooleanCellValue());

					break;

				default:

					// insertSql.append("'" + xssfCell.getRawValue() + "'");
					insertSql.append("''");
					// System.out.print(xssfCell.getRawValue());

				}

				break;

			case ERROR:

				// System.out.print(xssfCell.getErrorCellString());

				insertSql.append("''");
				// System.out.print("");

				break;

			default:

				insertSql.append("''");
				// insertSql.append("'" + xssfCell.getRawValue() + "'");
				// System.out.print(xssfCell.getRawValue());

			}

			if (cellNum != maxCellCount - 1) {
				insertSql.append(",");
			}
		}

		// Add a row delimiter between the output rows

		insertSql.append(")");

		log.warn(insertSql.toString());

		try {
			statement.execute(insertSql.toString());
		} catch (Exception e) {
			log.warn(e.toString());
		}
	}

	private List<String> getColNames(Row row, int maxCellCount) {
		List<String> colNames = new LinkedList<String>();

		for (int cellNum = 0; cellNum < maxCellCount; cellNum++) {

			// Get the cell

			Cell cell = row.getCell(cellNum);

			if (cell == null) {
				return new LinkedList<String>();
			}

			// Process the cell based on the cell
			// type

			switch (cell.getCellTypeEnum()) {

			case STRING:
				colNames.add(cell.getStringCellValue());

				break;

			default:

				return new LinkedList<String>();

			}

		}

		return colNames;

	}

	private int createWorksheetInformation(String insertSql) {
		Connection connection = null;
		Statement statement = null;

		int id = 0;
		try {
			connection = DriverManager.getConnection(mysqlUrl, mysqlUser, mysqlPassword);

			statement = connection.createStatement();

			statement.execute(insertSql);

			ResultSet resultSet = statement.executeQuery("select last_insert_id() as id");
			while (resultSet.next()) {
				id = resultSet.getInt(1);

				break;
			}

			resultSet.close();

			StringBuilder updateSql = new StringBuilder();
			updateSql.append("update sheetinformation set tablename = '");
			updateSql.append(tableNamePrefix + to10Numeric(id));
			updateSql.append("' where id = " + id);

			statement.executeUpdate(updateSql.toString());

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (statement != null) {
					statement.close();
				}
				if (connection != null) {
					connection.close();
				}
			} catch (Exception e) {

			}
		}

		return id;
	}

	private void createColumnInformation(String insertSql, Connection connection, Statement statement) {

		try {
			connection = DriverManager.getConnection(mysqlUrl, mysqlUser, mysqlPassword);

			statement = connection.createStatement();

			statement.execute(insertSql);

		} catch (Exception e) {
		} finally {
		}

	}

	private JSONObject resultSetToSheet(ResultSet resultSet) {
		JSONObject sheetObject = new JSONObject();

		try {

			int id = resultSet.getInt(1);
			sheetObject.put("id", id);

			String tableName = resultSet.getString(2);
			sheetObject.put("tableName", tableName);

			String sheetName = resultSet.getString(3);
			sheetObject.put("sheetName", sheetName);

			String alias = resultSet.getString(4);
			sheetObject.put("alias", alias);

			String description = resultSet.getString(5);
			sheetObject.put("description", description);

			String department = resultSet.getString(6);
			sheetObject.put("department", department);

			int records = resultSet.getInt(7);
			sheetObject.put("records", records);

			long timestamp = resultSet.getLong(8);
			sheetObject.put("timestamp", timestamp);
		} catch (Exception e) {
			log.warn(e.toString());
		}

		return sheetObject;
	}

	private JSONArray getColumnInformation(int sheetId, Connection connection, Statement statement) {

		JSONArray columinformationJsonArray = new JSONArray();

		try {

			String querySql = "select id, sheetinformationid, colname, nameinexcel from columninformation where sheetinformationid = "
					+ sheetId + " order by id asc";

			ResultSet resultSet = statement.executeQuery(querySql);
			while (resultSet.next()) {

				JSONObject columninformationJsonObject = resultSetToColumnInformation(resultSet);

				columinformationJsonArray.add(columninformationJsonObject);
			}

			resultSet.close();

		} catch (Exception e) {
			log.warn(e.toString());
		}

		return columinformationJsonArray;
	}

	private JSONObject resultSetToColumnInformation(ResultSet resultSet) {
		JSONObject columnObject = new JSONObject();

		try {

			int id = resultSet.getInt(1);
			columnObject.put("id", id);

			int sheetInformationId = resultSet.getInt(2);
			columnObject.put("sheetInformationId", sheetInformationId);

			String columnName = resultSet.getString(3);
			columnObject.put("columnName", columnName);

			String nameInExcel = resultSet.getString(4);
			columnObject.put("nameInExcel", nameInExcel);
		} catch (Exception e) {
			log.warn(e.toString());
		}

		return columnObject;

	}

	private JSONArray getRecords(int columns, String tableName, Connection connection, Statement statement) {
		JSONArray recordsJsonArray = new JSONArray();

		try {

			StringBuilder stringBuilder = new StringBuilder();

			stringBuilder.append("select ");

			for (int i = 0; i < columns; i++) {
				if (i == 0) {
					stringBuilder.append(" Col" + i);
				} else {
					stringBuilder.append(", Col" + i);
				}
			}

			stringBuilder.append(" from " + tableName);

			ResultSet resultSet = statement.executeQuery(stringBuilder.toString());

			while (resultSet.next()) {

				JSONArray recordJsonArray = new JSONArray();

				for (int i = 1; i <= columns; i++) {

					String value = resultSet.getString(i);

					recordJsonArray.add(value);

				}

				recordsJsonArray.add(recordJsonArray);

			}

			resultSet.close();

		} catch (Exception e) {
			log.warn(e.toString());
		}

		return recordsJsonArray;
	}

	private String getTableNameFromId(int id) {
		return tableNamePrefix + to10Numeric(id);
	}

	private String to10Numeric(int number) {
		StringBuilder result = new StringBuilder();
		result.append(number);
		while (result.length() < 10) {
			result.insert(0, "0");
		}

		return result.toString();
	}
}
