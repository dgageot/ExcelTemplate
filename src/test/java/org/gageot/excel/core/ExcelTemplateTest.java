package org.gageot.excel.core;

import static org.fest.assertions.Assertions.*;
import java.io.File;
import org.junit.Test;
import org.springframework.dao.DataAccessException;

public class ExcelTemplateTest {
	@Test(expected = IllegalArgumentException.class)
	public void springBeanFactory() {
		ExcelTemplate reader = new ExcelTemplate();

		reader.afterPropertiesSet();
	}

	@Test(expected = DataAccessException.class)
	public void unknownFile() {
		ExcelTemplate reader = new ExcelTemplate(new File("Z:/Unknown"));

		reader.read("Tests");
	}

	@Test
	public void readStringList() {
		ExcelTemplate reader = new ExcelTemplate("simple.xls", getClass());

		String[][] lines = reader.read("Tests");

		assertThat(lines).hasSize(3);
		assertThat(lines[0]).containsOnly("KEY1", "KEY2", "KEY3");
		assertThat(lines[1]).containsOnly("Value1", "Value2", "Value3");
		assertThat(lines[2]).containsOnly("Value10", "Value20", "Value30");
	}

	@Test
	public void readEmptySheet() {
		ExcelTemplate reader = new ExcelTemplate("empty.xls", getClass());

		String[][] lines = reader.read("Tests");

		assertThat(lines).isEmpty();
	}

	@Test
	public void readOneLine() {
		ExcelTemplate reader = new ExcelTemplate("oneLine.xls", getClass());

		String[][] lines = reader.read("Tests");

		assertThat(lines).hasSize(1);
		assertThat(lines[0]).containsOnly("KEY1", "KEY2", "KEY3");
	}

	@Test
	public void readEmptyLine() {
		ExcelTemplate reader = new ExcelTemplate("emptyLine.xls", getClass());

		String[][] lines = reader.read("Tests");

		assertThat(lines).hasSize(2);
		assertThat(lines[0]).containsOnly("KEY1", "KEY2", "KEY3");
		assertThat(lines[1]).containsOnly("A", "B", "C");
	}

	@Test
	public void readIndexedLines() {
		ExcelTemplate reader = new ExcelTemplate("indexedLines.xls", getClass());

		String[][] lines = reader.read("Tests");

		assertThat(lines).hasSize(3);
		assertThat(lines[0]).containsOnly("INDEX", "KEY1", "KEY2", "KEY3");
		assertThat(lines[1]).containsOnly("1");
		assertThat(lines[2]).containsOnly("2", "A", "B", "C");
	}
}
