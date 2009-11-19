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
import org.apache.poi.hssf.usermodel.HSSFCell;

/** 
 * An interface used by ExcelTemplate for mapping cells.
 * Implementations of this interface perform the actual work of mapping
 * cells, but don't need to worry about exception handling. IOExceptions
 * will be caught and handled correctly by the ExcelTemplate class.
 *
 * @author David Gageot
 * @see ExcelTemplate
 * @see SheetExtractor
 * @see RowMapper
 */
public interface CellMapper<T> {
	/** 
	 * Implementations must implement this method to map each cell of data
	 * in the HSSFSheet. This method should extract the values of the current cell.
	 * @param cell the HSSFCell to map
	 * @param rowNum the number of the current row
	 * @param columnNum the number of the current column
	 * @throws IOException if a IOException is encountered getting
	 * column values (that is, there's no need to catch IOException)
	 */
	T mapCell(HSSFCell cell, int rowNum, int columnNum) throws IOException;
}
