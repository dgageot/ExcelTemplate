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

package org.gageot.excel.beans;

import org.gageot.excel.core.BeanCellCallbackHandler;
import org.gageot.excel.core.CellCallbackHandler;
import org.gageot.excel.core.ExcelTemplate;
import org.springframework.beans.BeansException;

/**
 * Interface of ExcelTemplate's low-level JavaBeans infrastructure.
 * Provides operations to set property values.
 *
 * @author David Gageot
 * @see BeanCellCallbackHandler
 * @see ExcelTemplate#read(String,CellCallbackHandler)
 */
public interface BeanSetter {
	/**
	 * Set the value of a property on the current bean.
	 * @param bean bean instance to populate
	 * @param propertyName the name of the property (case insensitive)
	 * @param propertyValue the value of the property
	 */
	void setProperty(Object bean, String propertyName, Object propertyValue) throws BeansException;
}
