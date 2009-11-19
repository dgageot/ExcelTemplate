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
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.springframework.dao.DataAccessException;

/** 
 * Callback interface used by ExcelTemplate's query methods.
 * Implementations of this interface perform the actual work of extracting
 * results, but don't need to worry about exception handling. IOExceptions
 * will be caught and handled correctly by the ExcelTemplate class.
 *
 * <p>This interface is mainly used within the ExcelTemplate framework.
 * A RowCallbackHandler is usually a simpler choice for HSSFSheet processing.
 *
 * <p>Note: In contrast to a RowCallbackHandler, a SheetExtractor object
 * is typically stateless and thus reusable, as long as it doesn't access
 * stateful resources or keep result state within the object.
 *
 * @author David Gageot
 * @see ExcelTemplate
 * @see RowCallbackHandler
 */
public interface SheetExtractor<T> {
	/** 
	 * Implementations must implement this method to process
	 * all rows in the HSSFSheet.
	 * @param sheet HSSFSheet to extract data from.
	 * @return an arbitrary result object, or <code>null</code> if none
	 * (the extractor will typically be stateful in the latter case).
	 * @throws IOException if a IOException is encountered getting column
	 * values or navigating (that is, there's no need to catch IOException)
	 * @throws DataAccessException in case of custom exceptions
	 */
	T extractData(HSSFSheet sheet) throws IOException, DataAccessException;
}
