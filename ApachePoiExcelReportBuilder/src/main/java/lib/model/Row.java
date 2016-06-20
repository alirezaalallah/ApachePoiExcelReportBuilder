package lib.model;

import java.util.ArrayList;
import java.util.List;

public class Row {
    private List<Column> columns = new ArrayList<>();


    public Row addColumns(List<Column> columns) {
        this.columns.addAll(columns);
        return this;
    }

    public Row addColumn(Column column) {
        columns.add(column);
        return this;
    }

    public List<Column> columns() {
        return columns;
    }

    public static Row getNewInstance() {
        return new Row();
    }
}
