package org.gageot.excel.core;

import static org.fest.assertions.Assertions.*;
import java.util.List;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;
import org.springframework.beans.factory.BeanCreationException;

public class BeanCellCallbackHandlerTest {
	private static final String FILE_NAME = "beans.xls";
	private static final String TAB_NAME = "Tests";

	private ExcelTemplate reader;

	@Rule
	public ExpectedException expectedException = ExpectedException.none();

	@Before
	public void setUp() {
		reader = new ExcelTemplate(FILE_NAME, getClass());
	}

	@Test
	public void readBeans() {
		List<NameAndAge> beans = reader.readBeans(TAB_NAME, NameAndAge.class);

		assertThat(beans).hasSize(2);
		assertThat(beans.get(0).getLastName()).isEqualTo("Smith");
		assertThat(beans.get(0).getAge()).isEqualTo(35);
		assertThat(beans.get(1).getLastName()).isEqualTo("Johns");
		assertThat(beans.get(1).getAge()).isEqualTo(25);
	}

	@Test
	public void readBeansShouldIgnoreUnknownAttributes() {
		List<Name> beans = reader.readBeans(TAB_NAME, Name.class);

		assertThat(beans).hasSize(2);
		assertThat(beans.get(0).getLastName()).isEqualTo("Smith");
		assertThat(beans.get(1).getLastName()).isEqualTo("Johns");
	}

	@Test
	public void readBeanShouldFailWithWrongBeanType() {
		expectedException.expect(BeanCreationException.class);
		expectedException.expectMessage("Impossible to create bean");

		reader.readBeans(TAB_NAME, IllegalBean.class);
	}

	@Test
	public void readBeansShoudlFailWithPrivateBeanConstructor() {
		expectedException.expect(BeanCreationException.class);
		expectedException.expectMessage("Impossible to create bean");

		reader.readBeans(TAB_NAME, PrivateBean.class);
	}

	public static class NameAndAge {
		private String lastName;
		private double age;

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
	}

	public static class Name {
		private String lastName;

		public String getLastName() {
			return lastName;
		}

		public void setLastName(String lastName) {
			this.lastName = lastName;
		}
	}

	public static class PrivateBean {
		private PrivateBean() {
			// private constructor
		}
	}

	public static class IllegalBean {
		public IllegalBean(@SuppressWarnings("unused") int nop) {
			// private constructor
		}
	}
}
