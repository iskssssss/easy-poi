package top.kongsheng.common.easy_poi.handler;

import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.easy_poi.config.StyleConfig;
import top.kongsheng.common.easy_poi.enums.BorderStyle;
import top.kongsheng.common.easy_poi.enums.HorizontalAlignment;
import top.kongsheng.common.easy_poi.enums.VerticalAlignment;

import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

/**
 * 导出样式抽象类
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/9/13 10:59
 */
public abstract class AbsWriteStyle<CellStyleType> {

    /**
     * 数据是否换行
     */
    private boolean oneRow = false;

    /**
     * 字体
     */
    private String font = "宋体";
    /**
     * 字体大小
     */
    private float fontSizePoints = 11;
    /**
     * 是否粗体
     */
    private boolean fontBold = false;
    /**
     * 字体颜色
     */
    private String fontColor = "000000";

    HorizontalAlignment horizontal = HorizontalAlignment.LEFT;

    VerticalAlignment vertical = VerticalAlignment.CENTER;

    BorderStyle border = BorderStyle.NONE;
    BorderStyle borderTop = null;
    BorderStyle borderBottom = null;
    BorderStyle borderLeft = null;
    BorderStyle borderRight = null;

    /**
     * 行缩进
     */
    private int lineIndentNum;

    public abstract CellStyleType get();

    public boolean isOneRow() {
        return oneRow;
    }

    public AbsWriteStyle<CellStyleType> setOneRow(boolean oneRow) {
        this.oneRow = oneRow;
        return this;
    }

    public String getFont() {
        return font;
    }

    public AbsWriteStyle<CellStyleType> setFont(String font) {
        this.font = font;
        return this;
    }

    public float getFontSizePoints() {
        return fontSizePoints;
    }

    public AbsWriteStyle<CellStyleType> setFontSizePoints(float fontSizePoints) {
        this.fontSizePoints = fontSizePoints;
        return this;
    }

    public boolean isFontBold() {
        return fontBold;
    }

    public AbsWriteStyle<CellStyleType> setFontBold(boolean fontBold) {
        this.fontBold = fontBold;
        return this;
    }

    public String getFontColor() {
        return fontColor;
    }

    public AbsWriteStyle<CellStyleType> setFontColor(String fontColor) {
        this.fontColor = fontColor;
        return this;
    }

    public HorizontalAlignment getHorizontal() {
        return horizontal;
    }

    public AbsWriteStyle<CellStyleType> setHorizontal(HorizontalAlignment horizontal) {
        this.horizontal = horizontal;
        return this;
    }

    public AbsWriteStyle<CellStyleType> setHorizontal(String horizontal) {
        if (StrUtil.isEmpty(horizontal)) {
            return this;
        }
        this.horizontal = HorizontalAlignment.valueOf(horizontal.toUpperCase(Locale.ROOT));
        return this;
    }

    public VerticalAlignment getVertical() {
        return vertical;
    }

    public AbsWriteStyle<CellStyleType> setVertical(VerticalAlignment vertical) {
        this.vertical = vertical;
        return this;
    }

    public AbsWriteStyle<CellStyleType> setVertical(String vertical) {
        if (StrUtil.isEmpty(vertical)) {
            return this;
        }
        this.vertical = VerticalAlignment.valueOf(vertical.toUpperCase(Locale.ROOT));
        return this;
    }

    public BorderStyle getBorder() {
        return border;
    }

    public AbsWriteStyle<CellStyleType> setBorder(BorderStyle border) {
        this.border = border;
        return this;
    }

    public BorderStyle getBorderTop() {
        if (borderTop == null) {
            return this.border;
        }
        return borderTop;
    }

    public AbsWriteStyle<CellStyleType> setBorderTop(BorderStyle borderTop) {
        this.borderTop = borderTop;
        return this;
    }

    public BorderStyle getBorderBottom() {
        if (borderBottom == null) {
            return this.border;
        }
        return borderBottom;
    }

    public AbsWriteStyle<CellStyleType> setBorderBottom(BorderStyle borderBottom) {
        this.borderBottom = borderBottom;
        return this;
    }

    public BorderStyle getBorderLeft() {
        if (borderLeft == null) {
            return this.border;
        }
        return borderLeft;
    }

    public AbsWriteStyle<CellStyleType> setBorderLeft(BorderStyle borderLeft) {
        this.borderLeft = borderLeft;
        return this;
    }

    public BorderStyle getBorderRight() {
        if (borderRight == null) {
            return this.border;
        }
        return borderRight;
    }

    public AbsWriteStyle<CellStyleType> setBorderRight(BorderStyle borderRight) {
        this.borderRight = borderRight;
        return this;
    }

    public AbsWriteStyle<CellStyleType> setLineIndentNum(int lineIndentNum) {
        this.lineIndentNum = lineIndentNum;
        return this;
    }

    public int getLineIndentNum() {
        return lineIndentNum;
    }

    public abstract AbsWriteStyle<?> copy();

    public void init(StyleConfig titleStyleConfig) {
        if (titleStyleConfig == null) {
            return;
        }
        Integer border = titleStyleConfig.getBorder();
        BorderStyle borderStyle = BorderStyle.NONE;
        if (border != null && BORDER_STYLE_MAP.containsKey(border)) {
            borderStyle = BORDER_STYLE_MAP.get(border);
        }
        this.setFont(titleStyleConfig.getFont())
                .setFontBold(titleStyleConfig.isFontBold())
                .setFontColor(titleStyleConfig.getFontColor())
                .setFontSizePoints(titleStyleConfig.getFontSizePoints())
                .setHorizontal(titleStyleConfig.getHorizontal())
                .setVertical(titleStyleConfig.getVertical())
                .setBorder(borderStyle)
                .setLineIndentNum(titleStyleConfig.getLineIndentNum());
    }

    private static final Map<Integer, BorderStyle> BORDER_STYLE_MAP = new HashMap<>();

    static {
        BorderStyle[] values = BorderStyle.values();
        for (BorderStyle value : values) {
            BORDER_STYLE_MAP.put(((int) value.getCode()), value);
        }
    }
}
