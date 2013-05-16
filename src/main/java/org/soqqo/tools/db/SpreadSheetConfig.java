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

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.Resource;


/**
 * Simple immutable bean (Really, learn Scala!! case classes solve the world)
 * 
 * @author rbuckland
 * 
 */
public class SpreadSheetConfig {
	
	public enum ColumnType { 
		VARCHAR,
		NUMERIC,
		INT
	}

	private final Logger logger = LoggerFactory.getLogger(SpreadSheetConfig.class);
	
	// array of the names of the columns (the index is the column 0th, from the spreadsheet).
	private ArrayList<String> columnNames = new ArrayList<String>();
	
	// column types (crude way to rememeber what each type is)
	// keyed by the "column name"
	private HashMap<String,ColumnType> columnTypes = new HashMap<String,ColumnType>();

	/**
	 * Return the column type for the "index" number (the cell number in our SS).
	 * @param index
	 * @return
	 */
	public ColumnType getTypeForIndex(int index) {
		return columnTypes.get(columnNames.get(index));
	}
	public final Resource getSpreadSheetFile() {
		return spreadSheetFile;
	}

	public final void setSpreadSheetFile(Resource spreadSheetFile) {
		this.spreadSheetFile = spreadSheetFile;
	}

	public final HSSFSheet theSheet() {
		return sheet;
	}

	public final String getName() {
		return name;
	}

	public final void setName(String name) {
		this.name = name;
	}

	public final String getWorksheetName() {
		return worksheetName;
	}

	public final void setWorksheetName(String worksheetName) {
		this.worksheetName = worksheetName;
	}

	/**
	 * Helper loader to read the SS worksheet
	 * 
	 * @return
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public HSSFSheet loadSheet() {
		try {
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(spreadSheetFile.getFile()));
			HSSFWorkbook wb = new HSSFWorkbook(fs, true);
			if (!StringUtils.isEmpty(getWorksheetName())) {
				sheet = wb.getSheet(getWorksheetName());
				if (sheet == null) {
					throw new POIXLSDatabasePopulatorException("Could not find the Worksheet [" + getWorksheetName()
					        + "] from [" + spreadSheetFile.getDescription() + "]. Is that worksheet in there ?");
				}
			} else {
				sheet = wb.getSheetAt(0);
				if (sheet == null) {
					throw new POIXLSDatabasePopulatorException("Could not load a worksheet from ["
					        + spreadSheetFile.getDescription() + "]");
				}
			}
			return sheet;
		} catch (Exception e) {
			logger.error("Error loading " + spreadSheetFile.getDescription() + " : " + e.getLocalizedMessage());
			throw new RuntimeException(e);
		}

	}

	public ArrayList<String> getColumnNames() {
	    return columnNames;
    }

	public void setColumnNames(ArrayList<String> columnNames) {
	    this.columnNames = columnNames;
    }

	public HashMap<String,ColumnType> getColumnTypes() {
	    return columnTypes;
    }

	public void setColumnTypes(HashMap<String,ColumnType> columnTypes) {
	    this.columnTypes = columnTypes;
    }

	private String name;
	private Resource spreadSheetFile;
	private String worksheetName;
	private HSSFSheet sheet;
	
	public String toString() { 
		return "SpreadSheetConfig(" + name + "," + spreadSheetFile.getDescription() + "![" + worksheetName + "]," + columnNames + "," + columnTypes;
	}

}
