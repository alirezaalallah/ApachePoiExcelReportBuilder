package lib.model;

public class ExcelPoint {
    Integer columnIndex;
    Integer rowIndex;

    public ExcelPoint(Integer columnIndex, Integer rowIndex) {
        this.columnIndex = columnIndex;
        this.rowIndex = rowIndex;
    }

    public static ExcelPoint newPoint(Integer columnIndex, Integer rowIndex) {
        return new ExcelPoint(columnIndex, rowIndex);
    }

    public Integer getColumnIndex() {
        return columnIndex;
    }

    public void setColumnIndex(Integer columnIndex) {
        this.columnIndex = columnIndex;
    }

    public Integer getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(Integer rowIndex) {
        this.rowIndex = rowIndex;
    }
}
