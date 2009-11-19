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
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.util.LinkedCaseInsensitiveMap;

/**
 * RowMapper implementation that creates a <code>java.util.Map</code>
 * for each row, representing all columns as key-value pairs: one
 * entry for each column, with the column name as key.
 *
 * <p>The Map implementation to use can be customized through overriding
 * <code>createColumnMap</code>.
 *
 * <p>The CellMapper implementation to use can be customized through overriding
 * <code>createCellMapper</code>.
 *
 * @author David Gageot
 * @see ExcelTemplate#readList(String)
 */
public class ColumnMapRowMapper<T> implements RowMapper<Map<String, T>> {
	private final String[] keys;
	private final CellMapper<T> cellMapper;

	public ColumnMapRowMapper(String[] aKeys, CellMapper<T> cellMapper) {
		this.keys = aKeys;
		this.cellMapper = cellMapper;
	}

	public Map<String, T> mapRow(HSSFRow row, int rowNum) throws IOException {
		Map<String, T> map = createColumnMap(row.getLastCellNum());

		short lastColumnNum = row.getLastCellNum();
		for (short columnNum = 0; columnNum < lastColumnNum; columnNum++) {
			if (columnNum < keys.length) { // TODO : write test
				map.put(keys[columnNum], cellMapper.mapCell(row.getCell(columnNum, Row.RETURN_BLANK_AS_NULL), rowNum, columnNum));
			}
		}

		return map;
	}

	/**
	 * Create a Map instance to be used as column map.
	 * <p>By default, a linked case-insensitive Map will be created.
	 * @param columnCount the column count, to be used as initial
	 * capacity for the Map
	 * @return the new Map instance
	 * @see org.springframework.util.LinkedCaseInsensitiveMap
	 */
	protected Map<String, T> createColumnMap(int columnCount) {
		return new LinkedCaseInsensitiveMap<T>();
	}
}
