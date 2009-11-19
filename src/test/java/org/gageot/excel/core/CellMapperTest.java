package org.gageot.excel.core;

import static org.fest.assertions.Assertions.*;
import java.awt.Point;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.junit.Before;
import org.junit.Test;

public class CellMapperTest {
	private static final String FILE_NAME = "simple.xls";
	private static final String TAB_NAME = "Tests";

	private ExcelTemplate reader;

	@Before
	public void setUp() {
		reader = new ExcelTemplate(FILE_NAME, getClass());
	}

	@Test
	public void readObjectMatrix() {
		Point[][] lines = reader.read(TAB_NAME, new CellMapper<Point>() {
			public Point mapCell(HSSFCell cell, int rowNum, int columnNum) {
				return new Point(rowNum, columnNum);
			}
		}, Point.class);

		assertThat(lines).hasSize(3);

		for (int row = 0; row < 3; row++) {
			assertThat(lines[row]).hasSize(3);
			for (int col = 0; col < 3; col++) {
				assertThat(lines[row][col]).isEqualTo(new Point(row, col));
			}
		}
	}
}
