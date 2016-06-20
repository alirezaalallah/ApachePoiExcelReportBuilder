package lib.model;

import java.util.ArrayList;
import java.util.List;

public class Sheet {
    private String name;
    private Header header;
    private Boolean useMergeCell = true;
    private Boolean autoSizeColumn = true;
    private ExcelPoint startPoint = ExcelPoint.newPoint(0, 0);
    private Attribute direction = Attribute.DIRECTION_RTL;

    private List<Row> rows = new ArrayList<>();

    public Sheet addRows(List<Row> rows) {
        this.rows.addAll(rows);
        return this;
    }

    public Sheet addRow(Row rows) {
        this.rows.add(rows);
        return this;
    }

    public List<Row> rows() {
        return rows;
    }

    public Sheet direction(Attribute direction) {
        if (!direction.equals(Attribute.DIRECTION_LTR)
                && !direction.equals(Attribute.DIRECTION_RTL))
            throw new RuntimeException("invalid direction");

        this.direction = direction;
        return this;
    }

    public Attribute direction() {
        return direction;
    }

    public List<Row> rowsWithHeaderIfAvailable() {
        List<Row> rowsWithHeaderIfAvailable = new ArrayList<>();
        if (isHeaderAvailable()) {
            rowsWithHeaderIfAvailable.addAll(header.draw());
        }

        rowsWithHeaderIfAvailable.addAll(rows);

        return rowsWithHeaderIfAvailable;
    }

    public Sheet name(String name) {
        this.name = name;
        return this;
    }

    public Sheet autoSizeColumn(Boolean value) {
        this.autoSizeColumn = value;
        return this;
    }

    public Boolean autoSizeColumn() {
        return autoSizeColumn;
    }

    public String name() {
        return name;
    }

    public ExcelPoint startPoint() {
        return startPoint;
    }

    public Sheet startPoint(ExcelPoint startPoint) {
        this.startPoint = startPoint;
        return this;
    }

    public Sheet startPoint(Integer columnIndex, Integer rowIndex) {
        return startPoint(ExcelPoint.newPoint(columnIndex, rowIndex));
    }

    public static Sheet getNewInstance() {
        return new Sheet();
    }

    public Sheet header(Header header) {
        this.header = header;
        return this;
    }

    public Header header() {
        return header;
    }

    public Boolean isHeaderAvailable() {
        return header != null;
    }

}
