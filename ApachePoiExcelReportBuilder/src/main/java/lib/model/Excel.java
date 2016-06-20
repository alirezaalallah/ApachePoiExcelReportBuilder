package lib.model;

import java.util.ArrayList;
import java.util.List;

public class Excel {
    private List<Sheet> sheets = new ArrayList<>();

    public Excel addSheets(List<Sheet> sheets) {
        this.sheets.addAll(sheets);
        return this;
    }
    public Excel addSheet(Sheet sheets) {
        this.sheets.add(sheets);
        return this;
    }
    public List<Sheet> sheets() {
        return sheets;
    }

    public static Excel getNewInstance() {
        return new Excel();
    }

}
