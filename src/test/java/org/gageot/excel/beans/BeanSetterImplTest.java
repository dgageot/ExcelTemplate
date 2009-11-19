package org.gageot.excel.beans;

import static org.fest.assertions.Assertions.*;
import org.junit.Before;
import org.junit.Test;

public class BeanSetterImplTest {
	private MyBean bean;
	private BeanSetter beanSetter;

	@Before
	public void setUp() {
		bean = new MyBean();
		beanSetter = new BeanSetterImpl();
	}

	@Test
	public void setPropertyShouldIgnorePropertyCase() {
		beanSetter.setProperty(bean, "lastName", "Smith");
		assertThat(bean.getLastName()).isEqualTo("Smith");

		beanSetter.setProperty(bean, "lastname", "Johns");
		assertThat(bean.getLastName()).isEqualTo("Johns");

		beanSetter.setProperty(bean, "LASTNAME", "Jim");
		assertThat(bean.getLastName()).isEqualTo("Jim");
	}

	@Test
	public void setPropertyOnUnknownPropertyShouldntFail() {
		beanSetter.setProperty(bean, "unknownPropertyName", "Smith");
		// Shouldn't fail
	}

	@Test
	public void setPropertyWithConversionToDouble() {
		beanSetter.setProperty(bean, "age", "10");
		assertThat(bean.getAge()).isEqualTo(10.0);

		beanSetter.setProperty(bean, "age", new Integer(20));
		assertThat(bean.getAge()).isEqualTo(20.0);

		beanSetter.setProperty(bean, "age", new Double(30.0));
		assertThat(bean.getAge()).isEqualTo(30.0);
	}

	@Test
	public void setPropertyWithConversionToInteger() {
		beanSetter.setProperty(bean, "value", "10");
		assertThat(bean.getValue()).isEqualTo(10);

		beanSetter.setProperty(bean, "value", new Integer(20));
		assertThat(bean.getValue()).isEqualTo(20);

		beanSetter.setProperty(bean, "value", new Double(30.0));
		assertThat(bean.getValue()).isEqualTo(30);
	}

	public static class MyBean {
		private String lastName;
		private double age;
		private int value;

		public double getAge() {
			return age;
		}

		public void setAge(double age) {
			this.age = age;
		}

		public String getLastName() {
			return lastName;
		}

		public void setLastName(String lastName) {
			this.lastName = lastName;
		}

		public int getValue() {
			return value;
		}

		public void setValue(int value) {
			this.value = value;
		}
	}
}
