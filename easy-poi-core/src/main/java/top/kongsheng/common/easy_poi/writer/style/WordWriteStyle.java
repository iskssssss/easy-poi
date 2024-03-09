package top.kongsheng.common.easy_poi.writer.style;

import top.kongsheng.common.easy_poi.handler.AbsWriteStyle;

/**
 * word 样式
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/10/27 11:14
 */
public class WordWriteStyle extends AbsWriteStyle<WordWriteStyle> {

    @Override
    public WordWriteStyle get() {
        return this;
    }

    @Override
    public AbsWriteStyle<?> copy() {
        WordWriteStyle writeStyle = new WordWriteStyle();
        writeStyle.setOneRow(isOneRow());
        writeStyle.setFont(getFont());
        writeStyle.setFontSizePoints(getFontSizePoints());
        writeStyle.setFontBold(isFontBold());
        writeStyle.setFontColor(getFontColor());
        writeStyle.setHorizontal(getHorizontal());
        writeStyle.setVertical(getVertical());
        writeStyle.setBorder(getBorder());
        writeStyle.setBorderTop(getBorderTop());
        writeStyle.setBorderBottom(getBorderBottom());
        writeStyle.setBorderLeft(getBorderLeft());
        writeStyle.setBorderRight(getBorderRight());
        writeStyle.setLineIndentNum(getLineIndentNum());
        return writeStyle;
    }
}
