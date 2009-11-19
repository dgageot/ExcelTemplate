package org.gageot.excel.core;

import static org.fest.assertions.Assertions.*;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.junit.Before;
import org.junit.Test;

public class SheetExtractorTest {
	private static final String FILE_NAME = "simple.xls";
	private static final String TAB_NAME = "Tests";

	private ExcelTemplate reader;

	@Before
	public void initialize() {
		reader = new ExcelTemplate(FILE_NAME, getClass());
	}

	@Test
	public void readWithSheetExtractor() {
		int[] rows = reader.read(TAB_NAME, new SheetExtractor<int[]>() {
			public int[] extractData(HSSFSheet sheet) {
				return new int[] {
						sheet.getFirstRowNum(), sheet.getLastRowNum()
				};
			}
		});

		assertThat(rows[0]).isEqualTo(0);
		assertThat(rows[1]).isEqualTo(2);
	}
}
