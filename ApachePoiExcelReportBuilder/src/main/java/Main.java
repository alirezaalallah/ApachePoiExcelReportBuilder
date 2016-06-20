import lib.ApachePoiExcelBuilder;
import lib.model.*;

import java.util.ArrayList;
import java.util.List;

public class Main {

    public static void main(String[] args) {
        List<Row> rowsSheet = new ArrayList<>();

        for (Integer i = 0; i < 10; i++) {
            Row ibanRow = Row.getNewInstance();

            ibanRow.addColumn(Column.getNewInstance().value(1600000).index(0));
            ibanRow.addColumn(Column.getNewInstance().value("IR62000000000715318444").index(1));
            ibanRow.addColumn(Column.getNewInstance().value("IRR").index(2));
            ibanRow.addColumn(Column.getNewInstance().value("alireza").index(3));
            ibanRow.addColumn(Column.getNewInstance().value("alallah").index(4));

            rowsSheet.add(ibanRow);

        }

        Excel excel = Excel.getNewInstance()
                .addSheet(
                        Sheet.getNewInstance()
                                .name("test")
                                .autoSizeColumn(true)
                                .direction(Attribute.DIRECTION_LTR)
                                .startPoint(0, 0)
                                .addRows(rowsSheet)
                );
        try {

            ApachePoiExcelBuilder.getNewInstance()
                    .excel(excel)
                    .path("E:\\projects\\report\\new.xlsx")
                    .generate()
                    .save();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
