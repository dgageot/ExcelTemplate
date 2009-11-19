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

import java.beans.PropertyDescriptor;
import java.text.NumberFormat;
import org.gageot.excel.core.BeanCellCallbackHandler;
import org.gageot.excel.core.CellCallbackHandler;
import org.gageot.excel.core.ExcelTemplate;
import org.springframework.beans.BeanWrapper;
import org.springframework.beans.BeansException;
import org.springframework.beans.PropertyAccessorFactory;
import org.springframework.beans.propertyeditors.CustomNumberEditor;

/**
 * Simple BeanSetter implementation based on BeanWrapperImpl Spring class.
 *
 * @author David Gageot
 * @see BeanCellCallbackHandler
 * @see ExcelTemplate#read(String,CellCallbackHandler)
 */
public class BeanSetterImpl implements BeanSetter {
	private BeanWrapper currentWrapper;
	private Object currentBean;

	protected BeanWrapper getWrapper(Object bean) {
		if (bean != currentBean) {
			currentBean = bean;
			currentWrapper = PropertyAccessorFactory.forBeanPropertyAccess(currentBean);

			// To make sure that Double values can be converted into an int.
			//
			currentWrapper.registerCustomEditor(int.class, new CustomNumberEditor(Integer.class, NumberFormat.getInstance(), false));
		}

		return currentWrapper;
	}

	/**
	 * Set the value of a property on the current bean.
	 * 
	 * @param bean bean instance to populate
	 * @param propertyName the name of the property (case insensitive)
	 * @param propertyValue the value of the property
	 */
	public void setProperty(Object bean, String propertyName, Object propertyValue) throws BeansException {
		BeanWrapper wrapper = getWrapper(bean);

		PropertyDescriptor[] propertyDescriptors = wrapper.getPropertyDescriptors();

		// Find a bean property by its name ignoring case.
		// this way, we can accept any type of case in the spreadsheet file
		//
		for (PropertyDescriptor propertyDescriptor : propertyDescriptors) {
			if (propertyDescriptor.getName().equalsIgnoreCase(propertyName)) {
				wrapper.setPropertyValue(propertyDescriptor.getName(), propertyValue.toString());
				break;
			}
		}
	}
}
