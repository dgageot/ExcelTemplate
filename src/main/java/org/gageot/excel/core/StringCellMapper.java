/*
 * Copyright 2002-2009 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.gageot.excel.core;

import java.io.IOException;
import java.text.SimpleDateFormat;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;

/**
 * CellMapper implementation that creates a <code>java.lang.String</code>
 * for each cell.
 *
 * @author David Gageot
 */
public class StringCellMapper implements CellMapper<String> {
	private static final int TEXT_CELL_FORMAT = 49;
	private static final int OPENOFFICE_TEXT_CELL_FORMAT = 165;
	private static final int OPENOFFICE_DATE_CELL_FORMAT = 167;

	private SimpleDateFormat dateFormat;

	@Override
	public String mapCell(HSSFCell cell, int rowNum, int columnNum) throws IOException {
		if (null == cell) {
			return ""; // TODO
		}

		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				return "";
			case Cell.CELL_TYPE_ERROR:
				return "Error<" + cell.getErrorCellValue() + ">";
			case Cell.CELL_TYPE_BOOLEAN:
				return booleanToString(cell);
			case Cell.CELL_TYPE_NUMERIC:
				return numericToString(cell);
			case Cell.CELL_TYPE_FORMULA:
				return formulaToString(cell);
			case Cell.CELL_TYPE_STRING:
			default:
				return richTextToString(cell);
		}
	}

	private String booleanToString(HSSFCell cell) {
		return cell.getBooleanCellValue() ? "VRAI" : "FAUX";
	}

	private String richTextToString(HSSFCell cell) {
		return cell.getStringCellValue();
	}

	private String numericToString(HSSFCell cell) {
		double numericValue = cell.getNumericCellValue();

		if (Double.isNaN(numericValue)) {
			return "";
		}

		if (isDateFormat(cell)) {
			if (null == dateFormat) {
				dateFormat = new SimpleDateFormat("dd/MM/yyyy");
			}

			return dateFormat.format(cell.getDateCellValue());
		}

		// For text cells, Excel still tries to converts the content into
		// numerical value. For integer content, we want to convert
		// into a String value without fraction.
		//
		if (isTextFormat(cell) && (((long) numericValue) == numericValue)) {
			return Long.toString((long) numericValue);
		}

		return Double.toString(numericValue);
	}

	private String formulaToString(HSSFCell cell) {
		if (isTextFormat(cell)) {
			return richTextToString(cell);
		}

		return numericToString(cell);
	}

	private static boolean isTextFormat(HSSFCell cell) {
		short cellFormat = cell.getCellStyle().getDataFormat();

		return ((TEXT_CELL_FORMAT == cellFormat) || (OPENOFFICE_TEXT_CELL_FORMAT == cellFormat));
	}

	private static boolean isDateFormat(HSSFCell cell) {
		short cellFormat = cell.getCellStyle().getDataFormat();

		return (OPENOFFICE_DATE_CELL_FORMAT == cellFormat);
	}
}
