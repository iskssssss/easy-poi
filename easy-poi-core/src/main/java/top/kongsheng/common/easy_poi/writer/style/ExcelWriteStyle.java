package top.kongsheng.common.easy_poi.writer.style;

import top.kongsheng.common.easy_poi.handler.AbsWriteStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * excel 导出样式
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/10/26 17:03
 */
public class ExcelWriteStyle extends AbsWriteStyle<CellStyle> {

    private final Workbook workbook;
    private CellStyle cellStyle;

    public ExcelWriteStyle(Workbook workbook) {
        this.workbook = workbook;
    }

    @Override
    public CellStyle get() {
        if (this.cellStyle == null) {
            this.to();
        }
        return this.cellStyle;
    }

    @Override
    public AbsWriteStyle<?> copy() {
        ExcelWriteStyle dataStyle = new ExcelWriteStyle(this.workbook);
        dataStyle.setFont(getFont()).setFontBold(isFontBold()).setFontColor(getFontColor()).setFontSizePoints(getFontSizePoints())
                .setHorizontal(getHorizontal()).setVertical(getVertical())
                .setBorder(getBorder());
        return dataStyle;
    }

    private void to() {
        this.cellStyle = workbook.createCellStyle();
        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName(super.getFont());
        java.awt.Color awtColor = new java.awt.Color(Integer.valueOf(super.getFontColor(), 16));
        XSSFColor xssfColor = new XSSFColor(awtColor, null);
        font.setColor(xssfColor);
        font.setBold(super.isFontBold());
        font.setFontHeight((short) (super.getFontSizePoints() * 20));
        cellStyle.setAlignment(alignment(super.getHorizontal()));
        cellStyle.setVerticalAlignment(alignment(super.getVertical()));
        cellStyle.setFont(font);
        cellStyle.setBorderTop(borderStyle(super.getBorderTop()));
        cellStyle.setBorderBottom(borderStyle(super.getBorderBottom()));
        cellStyle.setBorderLeft(borderStyle(super.getBorderLeft()));
        cellStyle.setBorderRight(borderStyle(super.getBorderRight()));
        cellStyle.setWrapText(true);
    }

    public HorizontalAlignment alignment(top.kongsheng.common.easy_poi.enums.HorizontalAlignment source) {
        return HorizontalAlignment.forInt(source.getCode());
    }

    public VerticalAlignment alignment(top.kongsheng.common.easy_poi.enums.VerticalAlignment source) {
        if (source == null) {
            return null;
        }
        return VerticalAlignment.forInt(source.getCode());
    }

    public BorderStyle borderStyle(top.kongsheng.common.easy_poi.enums.BorderStyle source) {
        return BorderStyle.valueOf(source.getCode());
    }

    public Workbook getWorkbook() {
        return workbook;
    }
}
