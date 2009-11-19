package org.gageot.excel.core;

import static org.fest.assertions.Assertions.*;
import org.junit.Before;
import org.junit.Test;

public class StringCellMapperTest {
	private static final String FILE_NAME = "cellFormat.xls";
	private static final String TAB_NAME = "Tests";

	private ExcelTemplate reader;

	@Before
	public void initialize() {
		reader = new ExcelTemplate(FILE_NAME, getClass());
	}

	@Test
	public void readWithDefaultStringCellMapper() {
		String[][] lines = reader.read(TAB_NAME);

		assertThat(lines).hasSize(2);
		assertThat(lines[1]).containsOnly("1.0", "A", "3.0", "1.5", "1,5", "1", "VRAI", "FAUX", "", "31/01/2007", "Text Formula");
	}

	@Test
	public void readWithExplicitStringCellMapper() {
		String[][] lines = reader.read(TAB_NAME, new StringCellMapper(), String.class);

		assertThat(lines).hasSize(2);
		assertThat(lines[1]).containsOnly("1.0", "A", "3.0", "1.5", "1,5", "1", "VRAI", "FAUX", "", "31/01/2007", "Text Formula");
	}
}