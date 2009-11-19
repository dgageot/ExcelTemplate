/*
 * Copyright 2002-2005 the original author or authors.
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
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import com.google.common.collect.Lists;

/**
 * Adapter implementation of the SheetExtractor interface that delegates
 * to a RowMapper which is supposed to create an object for each row.
 * Each object is added to the results List of this SheetExtractor.
 *
 * <p>Useful for the typical case of one object per row in the Excel spreadsheet.
 * The number of entries in the results list will match the number of rows.
 *
 * <p>Note that a RowMapper object is typically stateless and thus reusable;
 * just the RowMapperResultSetExtractor adapter is stateful.
 *
 * @author David Gageot
 * @see RowMapper
 */
public class RowMapperSheetExtractor<T> implements SheetExtractor<List<T>> {
	private final RowMapper<T> rowMapper;

	/**
	 * Create a new RowMapperSheetExtractor.
	 * @param aRowMapper the RowMapper which creates an object for each row
	 */
	public RowMapperSheetExtractor(RowMapper<T> rowMapper) {
		this.rowMapper = rowMapper;
	}

	public List<T> extractData(HSSFSheet sheet) throws IOException {
		List<T> rows = Lists.newArrayList();

		int firstRowIndex = sheet.getFirstRowNum();
		int lastRowIndex = sheet.getLastRowNum();

		for (int i = firstRowIndex; i <= lastRowIndex; i++) {
			T row = rowMapper.mapRow(sheet.getRow(i), i);
			if (null != row) {
				rows.add(row);
			}
		}

		return rows;
	}
}
