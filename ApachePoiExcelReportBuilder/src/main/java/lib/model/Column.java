package lib.model;

import java.math.BigInteger;

public class Column {
    private String imagePath;
    private String value;
    private String formula;
    private Integer colSpan;
    private Integer index;
    private Type type = Type.CELL_TYPE_STRING;
    private Attribute style = Attribute.STYLE_NONE;
    private Attribute alignment = Attribute.ALIGN_CENTER;


    public Column style(Attribute style) {
        this.style = style;
        return this;
    }

    public Attribute style() {
        return style;
    }

    public Column type(Type type) {
        this.type = type;
        return this;
    }

    public Column alignment(Attribute align) {
        this.alignment = align;
        return this;
    }

    public Attribute alignment() {
        return this.alignment;
    }

    public Short alignmentIndex() {
        Short alignmentIndex = 0;
        switch (alignment) {
            case ALIGN_GENERAL:
                alignmentIndex = 0;
                break;
            case ALIGN_LEFT:
                alignmentIndex = 1;
                break;
            case ALIGN_CENTER:
                alignmentIndex = 2;
                break;
            case ALIGN_RIGHT:
                alignmentIndex = 3;
                break;
            case ALIGN_FILL:
                alignmentIndex = 4;
                break;
            case ALIGN_JUSTIFY:
                alignmentIndex = 5;
                break;
            case ALIGN_CENTER_SELECTION:
                alignmentIndex = 6;
                break;
            case ALIGN_VERTICAL_TOP:
                alignmentIndex = 0;
                break;
            case ALIGN_VERTICAL_CENTER:
                alignmentIndex = 1;
                break;
            case ALIGN_VERTICAL_BOTTOM:
                alignmentIndex = 2;
                break;
            case ALIGN_VERTICAL_JUSTIFY:
                alignmentIndex = 3;
                break;
        }

        return alignmentIndex;
    }

    public Type type() {
        return type;
    }

    public Integer index() {
        return index;
    }

    public Column index(Integer index) {
        this.index = index;
        return this;
    }

    public Integer colSpan() {
        return colSpan;
    }

    public Column colSpan(Integer colSpan) {
        this.colSpan = colSpan;
        return this;
    }

    public String formula() {
        return formula;
    }

    public Column formula(String formula) {
        this.formula = formula;
        return this;
    }

    public String value() {
        return value;
    }

    public Column value(String value) {
        type(Type.CELL_TYPE_STRING);
        this.value = value;
        return this;
    }

    public Column value(BigInteger value) {
        type(Type.CELL_TYPE_NUMERIC);
        this.value = value.toString();
        return this;
    }

    public Column value(Double value) {
        type(Type.CELL_TYPE_NUMERIC);
        this.value = value.toString();
        return this;
    }

    public Column value(Long value) {
        type(Type.CELL_TYPE_NUMERIC);
        this.value = value.toString();
        return this;
    }

    public Column value(Float value) {
        type(Type.CELL_TYPE_NUMERIC);
        this.value = value.toString();
        return this;
    }

    public Column value(Integer value) {
        type(Type.CELL_TYPE_NUMERIC);
        this.value = value.toString();
        return this;
    }

    public Column imagePath(String imagePath) {
        this.imagePath = imagePath;
        return this;
    }

    public String imagePath() {
        return imagePath;
    }

    public static Column getNewInstance() {
        return new Column();
    }

    public enum Type {
        CELL_TYPE_NUMERIC(0),
        CELL_TYPE_STRING(1),
        CELL_TYPE_FORMULA(2),
        CELL_TYPE_BLANK(3),
        CELL_TYPE_BOOLEAN(4),
        CELL_TYPE_ERROR(5);

        private Integer type;

        private Type(Integer type) {
            this.type = type;
        }

        public Integer getType() {
            return type;
        }
    }
}
