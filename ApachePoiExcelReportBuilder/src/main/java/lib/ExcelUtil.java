package lib;


import lib.model.Column;
import lib.model.Row;

import java.util.List;

public class ExcelUtil {
    private static final ExcelUtil instance = new ExcelUtil();

    private ExcelUtil() {
    }

    public static ExcelUtil getInstance() {
        return instance;
    }

    public Integer getMaxColumnIndex(List<Row> rows) {
        Integer lastColumn = 0;
        for (Row row : rows) {

            Integer maxColumn = 0;
            for (Column column : row.columns()) {
                if (column.colSpan() != null) {
                    maxColumn += column.colSpan();
                } else {
                    maxColumn++;
                }
            }

            if (maxColumn > lastColumn) {

                lastColumn = maxColumn;
            }

        }
        return lastColumn;
    }
}
