package org.gageot.excel.core;

import static org.fest.assertions.Assertions.*;
import static org.fest.assertions.MapAssert.*;
import java.util.List;
import java.util.Map;
import org.junit.Before;
import org.junit.Test;

public class ColumnMapRowMapperTest {
	private static final String FILE_NAME = "simple.xls";
	private static final String TAB_NAME = "Tests";

	private ExcelTemplate reader;

	@Before
	public void initialize() {
		reader = new ExcelTemplate(FILE_NAME, getClass());
	}

	@Test
	public void readListWithExplicitColumnMapRowMapper() {
		List<Map<String, String>> lines = reader.readList(TAB_NAME, new ColumnMapRowMapper<String>(new String[] {
				"KEY1", "KEY2", "KEY3"
		}, new StringCellMapper()));

		assertThat(lines).hasSize(3);
		assertThat(lines.get(0)).hasSize(3).includes(entry("KEY1", "KEY1")).includes(entry("KEY2", "KEY2")).includes(entry("KEY3", "KEY3"));
		assertThat(lines.get(1)).hasSize(3).includes(entry("KEY1", "Value1")).includes(entry("KEY2", "Value2")).includes(entry("KEY3", "Value3"));
		assertThat(lines.get(2)).hasSize(3).includes(entry("KEY1", "Value10")).includes(entry("KEY2", "Value20")).includes(entry("KEY3", "Value30"));
	}

	@Test
	public void readListWithImplicitColumnMapRowMapper() {
		List<Map<String, String>> lines = reader.readList(TAB_NAME);

		assertThat(lines).hasSize(2);
		assertThat(lines.get(0)).hasSize(3).includes(entry("KEY1", "Value1")).includes(entry("KEY2", "Value2")).includes(entry("KEY3", "Value3"));
		assertThat(lines.get(1)).hasSize(3).includes(entry("KEY1", "Value10")).includes(entry("KEY2", "Value20")).includes(entry("KEY3", "Value30"));
	}
}
