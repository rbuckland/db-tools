/*

   Copyright 2013 Soqqo Limited

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
  
 */
package org.soqqo.tools.db;

import java.sql.Connection;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.datasource.init.DatabasePopulator;

/**
 * This is a simple class that is configured with a map of classpath resource
 * XLS files, a worksheet name (omit if the only sheet) and an optional cell
 * reference
 * 
 * From here the data found in the SS will be loaded into a the datasource ..
 * one SQL table per sheet configured.
 * 
 * @author rbuckland
 * 
 */
public class POIXLSDatabasePopulator implements DatabasePopulator {

	private final String STRING_COLUMN = "VARCHAR(2048)";
	private final String NUMERIC_COLUMN = "NUMERIC(20,10)";

	private final Logger logger = LoggerFactory.getLogger(POIXLSDatabasePopulator.class);

	// name (will derive the table), filename resource, worksheet name
	private List<SpreadSheetConfig> spreadsheets;

	@Override
	public void populate(Connection connection) throws SQLException {
		for (SpreadSheetConfig config : spreadsheets) {
			config.loadSheet();
		}

		createTables(connection);

		insertData(connection);

	}

	private void insertData(Connection connection) throws SQLException {
		// start the SQL STAtement
		StringBuffer sb = new StringBuffer();

		for (SpreadSheetConfig config : spreadsheets) {

			sb.append("INSERT INTO ").append(config.getName()).append(" VALUES ");

			// starting at the "second row"
			for (int rn = 1; rn < config.theSheet().getLastRowNum(); rn++) {

				if (rn > 1) { 
					sb.append(",");
				}
				// start in the sheet
				Row r = config.theSheet().getRow(rn);

				// get the last column
				int lastColumn = r.getLastCellNum();

				sb.append("\n( ");

				// from first to last column
				for (int cn = 0; cn < lastColumn; cn++) {
					if (cn > 0) sb.append(", ");
					if(r.getCell(cn) == null) { 
						sb.append("null");
					} else { 
						switch (config.getTypeForIndex(cn)) {
						  case NUMERIC: sb.append(r.getCell(cn).getNumericCellValue()); break;
						  // TODO throw exception when the column name is bad
						  case VARCHAR: sb.append("'").append(getValueFromCell(r.getCell(cn))).append("'"); break;
						}
					}
				}
				sb.append(" )");
				
			}
			sb.append(";\n");
		}
		
		
		// executre the row
		String nativeSql = connection.nativeSQL(sb.toString());
		logger.debug("Executing " + nativeSql);
		Statement stmt = connection.createStatement();
		stmt.execute(nativeSql);
		stmt.close();

	}

	/**
	 * For each "sheet" we have loaded, we will create a table based on what we
	 * find in the sheet
	 * 
	 * First row with data, first column expects header names, these will be the
	 * column names.
	 * 
	 * @param connection
	 * @param sheets
	 * @throws SQLException
	 */
	private void createTables(Connection connection) throws SQLException {
		for (SpreadSheetConfig config : spreadsheets) {

			StringBuffer sb = new StringBuffer();
			sb.append("CREATE TABLE ").append(config.getName()).append(" ( ");
			sb.append(deriveHeaders(config));
			sb.append(" );");

			String nativeSql = connection.nativeSQL(sb.toString());
			logger.debug("Executing " + nativeSql);
			Statement stmt = connection.createStatement();
			stmt.execute(nativeSql);
			stmt.close();

		}
	}

	private String deriveHeaders(SpreadSheetConfig config) {
		StringBuffer sb = new StringBuffer();

		int firstRow = config.theSheet().getFirstRowNum();
		Row r = config.theSheet().getRow(firstRow);

		int lastColumn = r.getLastCellNum();

		for (int cn = 0; cn < lastColumn; cn++) {
			if (cn > 0)
				sb.append(", ");
			Cell c = r.getCell(cn, Row.RETURN_BLANK_AS_NULL);
			if (c == null) {
				throw new POIXLSDatabasePopulatorException("Sheet [" + config.getSpreadSheetFile().getDescription()
				        + "!" + config.getWorksheetName() + "] Row " + firstRow + ", Col " + cn
				        + " is missing a header for column name, I can't continue with this sheet");
			} else {
				sb.append(getColumnNameIDentifier(c, config)).append(" ").append(determinColumnType(c, config));
			}
		}

		return sb.toString();
	}
	
	/**
	 * TODO use the DataFormatter
	 * @param c
	 * @return
	 */
	private String getValueFromCell(Cell c) {
		switch (c.getCellType()) {
	      case Cell.CELL_TYPE_STRING:
	        return c.getStringCellValue();
		  case Cell.CELL_TYPE_BOOLEAN: 
			 return Boolean.toString(c.getBooleanCellValue());
		  case Cell.CELL_TYPE_NUMERIC: 
             return Double.toString(c.getNumericCellValue());
		  case Cell.CELL_TYPE_ERROR:
			 return String.valueOf(c.getErrorCellValue());
		  case Cell.CELL_TYPE_FORMULA:
			 return String.valueOf(c.getCellFormula());
		  case Cell.CELL_TYPE_BLANK:
		     return "";
	      default: throw new POIXLSDatabasePopulatorException("Can't determine the cell value");
    	}
	}

	/**
	 * Given an Excel cell positioned at the first row in the datatable, derive
	 * the column name for the SQL CREATE TABLE statement and the DATATYPE
	 * TODO throw an exception when the column name is not an SQL regexp
	 */
	private String getColumnNameIDentifier(Cell c, SpreadSheetConfig config) {
		String colName = "";
		String errorReColumn = "";
		switch (c.getCellType()) {
			case Cell.CELL_TYPE_STRING: {
				colName = c.getStringCellValue();
				config.getColumnNames().add(c.getColumnIndex(), colName);
				return colName;
			}
			case Cell.CELL_TYPE_BOOLEAN: errorReColumn = "We can't use a BOOLEAN value for a column name";
			case Cell.CELL_TYPE_NUMERIC: errorReColumn = "We can't use a NUMERIC value for a column name";
			case Cell.CELL_TYPE_ERROR: errorReColumn = "Cell has an error, can't use it for a column name";
			case Cell.CELL_TYPE_FORMULA: errorReColumn = "Cannot use a formula for the Column Name";
			case Cell.CELL_TYPE_BLANK: errorReColumn = "We need a column name (it's blank)";
			default:
				throw new POIXLSDatabasePopulatorException(
						"[" + config.getSpreadSheetFile().getDescription() + "]" +
				        new org.apache.poi.ss.util.CellReference(c).formatAsString() + " "
				        + errorReColumn);
		}
	}

	/**
	 * This method hunts down each row for the cell provided (the next row) and
	 * tries to determine what the "data" type of that column is. It first
	 * assumes a number, but it if finds anything other that that (like a string
	 * looking) thing, then it will quickly exit to a string.
	 * 
	 * TODO determine a better SQL type given the data found. (BOOLEAN, IN, and
	 * accurate VARCHAR size etc)
	 * 
	 * @param c
	 * @param config
	 * @return
	 */
	private String determinColumnType(Cell c, SpreadSheetConfig config) {
		Sheet sheet = c.getSheet();
		// we will assume that the column is all numeric, so if we find anything
		// else, then it's a string column
		for (int rowIdx = c.getRowIndex() + 1; rowIdx < c.getSheet().getLastRowNum(); rowIdx++) {
			// logger.debug("trying Cell(" + rowIdx + "," + c.getColumnIndex() +
			// ")");
			Cell cNext = sheet.getRow(rowIdx).getCell(c.getColumnIndex());
			if (cNext == null || cNext.getCellType() != Cell.CELL_TYPE_NUMERIC) {
				//logger.debug("adding[" + c.getColumnIndex() + "]," + config.getColumnNames().get(c.getColumnIndex()) + "," + SpreadSheetConfig.ColumnType.VARCHAR);
				config.getColumnTypes().put(config.getColumnNames().get(c.getColumnIndex()),SpreadSheetConfig.ColumnType.VARCHAR);
				return STRING_COLUMN + " NULL";
			} // TODO . determine the oher types .. rewrite this .. 
			// TODO Integer check - dbl == Math.rint(dbl)
		}
		//logger.debug("Cell(?," + c.getColumnIndex() + ") is Numeric");
		config.getColumnTypes().put(config.getColumnNames().get(c.getColumnIndex()),SpreadSheetConfig.ColumnType.NUMERIC);
		//logger.debug("adding[" + c.getColumnIndex() + "]," + config.getColumnNames().get(c.getColumnIndex()) + "," + SpreadSheetConfig.ColumnType.NUMERIC);
		return NUMERIC_COLUMN + " NULL";
	}

	public List<SpreadSheetConfig> getSpreadsheets() {
		return spreadsheets;
	}

	public void setSpreadsheets(List<SpreadSheetConfig> spreadsheets) {
		this.spreadsheets = spreadsheets;
	}

}
