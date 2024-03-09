package top.kongsheng.common.easy_poi.model;

import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.handler.AbsWriteStyle;
import top.kongsheng.common.easy_poi.utils.SizeConvertUtil;

import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * Cell
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/9/15 14:05
 */
public abstract class Cell<T> {
    private T cell;
    private AbsWriteStyle<?> writeStyle;
    private List<DrawFieldInfo> drawFieldInfoList;
    private Map<PoiModelField, DrawFieldInfo> drawFieldInfoMap;

    /**
     * 初始化样式信息
     *
     * @param writeStyle 样式信息
     */
    public void initStyle(AbsWriteStyle<?> writeStyle) {
        this.writeStyle = writeStyle;
    }

    /**
     * 写入列表数据
     *
     * @param poiItem   poi注解
     * @param cellValue 值
     */
    public abstract void writeList(PoiModelField poiItem, List<?> cellValue);

    /**
     * 写入字符串信息
     *
     * @param poiItem   poi注解
     * @param cellValue 值
     */
    public abstract void writeString(PoiModelField poiItem, String cellValue);

    /**
     * 设置背景颜色
     *
     * @param color 颜色
     */
    public abstract void setBackgroundColor(String color);

    /**
     * 设置字体颜色
     *
     * @param color 颜色
     */
    public abstract void setFontColor(String color);

    protected int getLineIndentNum(String cellValue) {
        int lineIndentNum = writeStyle.getLineIndentNum();
        if (lineIndentNum == -1) {
            List<DrawFieldInfo> drawFieldInfoList = getDrawFieldInfoList();
            final DrawFieldInfo drawFieldInfo = drawFieldInfoList.iterator().next();
            float cellRowMaxCharCount = drawFieldInfo.getCellRowMaxCharCount();
            float contentRow = SizeConvertUtil.getCellContentRow(cellValue, cellRowMaxCharCount);
            if (contentRow > 1) {
                lineIndentNum = 4;
            }
        }
        return lineIndentNum;
    }

    public T getCell() {
        return cell;
    }

    public void setCell(T cell) {
        this.cell = cell;
    }

    public AbsWriteStyle<?> getWriteStyle() {
        return writeStyle;
    }

    public void setWriteStyle(AbsWriteStyle<?> writeStyle) {
        this.writeStyle = writeStyle;
    }

    public List<DrawFieldInfo> getDrawFieldInfoList() {
        return drawFieldInfoList;
    }

    public void setDrawFieldInfoList(List<DrawFieldInfo> drawFieldInfoList) {
        this.drawFieldInfoList = drawFieldInfoList;
        if (this.drawFieldInfoList == null) {
            return;
        }
        this.drawFieldInfoMap = this.drawFieldInfoList.stream().collect(Collectors.toMap(DrawFieldInfo::getDrawPoiModelField, item -> item));
    }

    public Map<PoiModelField, DrawFieldInfo> getDrawFieldInfoMap() {
        return drawFieldInfoMap;
    }
}
