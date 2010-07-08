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
 * CellMapper implementation that creates a <code>java.lang.Object</code>
 * for each cell. It uses <code>java.lang.String</code> for text cells and
 * <code>java.lang.Double</code> for numerical cells.
 *
 * @author David Gageot
 */
public class ObjectCellMapper implements CellMapper<Object> {
	@Override
	public Object mapCell(HSSFCell cell, int rowNum, int columnNum) throws IOException {
		try {
			return cell.getNumericCellValue();
		} catch (NumberFormatException e) {
			return cell.getStringCellValue();
		} catch (IllegalStateException e) {
			return cell.getStringCellValue();
		}
	}
}
