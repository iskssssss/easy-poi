package top.kongsheng.common.easy_poi.model;

import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.handler.AbsWriteStyle;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.ToString;

import java.lang.reflect.Field;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Supplier;

/**
 * 绘制字段信息
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/10/11 17:25
 */
@Data
@ToString
@EqualsAndHashCode(callSuper = true)
public class DrawFieldInfo extends PoiModelFieldWriteConfig {

    /**
     * 是否是自定义字段
     */
    private boolean custom = false;
    /**
     * 预设横坐标
     */
    private int sourceX;
    /**
     * 实际绘制横坐标
     */
    private int drawX;
    /**
     * 同一单元格时的绘制顺序
     */
    private int drawSort;
    /**
     * 写入类型 1.覆盖 3.插入 2.追加
     */
    private int writeType;

    private boolean nested;
    /**
     * 是否换行绘制
     */
    private boolean newLine;
    /**
     * 宽度
     */
    private float width;
    /**
     * 绘制字段
     */
    private Field drawField;
    /**
     * 绘制字段配置信息
     */
    private PoiModelField drawPoiModelField;
    /**
     * 标题单元样式
     */
    private AbsWriteStyle<?> titleColStyle;
    /**
     * 数据单元样式
     */
    private AbsWriteStyle<?> dataColStyle;
    /**
     * 一行最大可写字符数量
     */
    private float cellRowMaxCharCount;

    private List<DrawFieldInfo> child;

    private Map<Long, AbsWriteStyle<?>> createDataCelStyleMap = new LinkedHashMap<>();

    /**
     * 标题单元样式
     *
     * @return 标题单元样式
     */
    public AbsWriteStyle<?> createTitleColStyle(Supplier<AbsWriteStyle<?>> createWriteStyle) {
        if (this.titleColStyle == null) {
            this.titleColStyle = createWriteStyle.get();
            this.titleColStyle.init(getTitleStyleConfig());
        }
        return this.titleColStyle;
    }

    /**
     * 数据单元样式
     *
     * @return 获取
     */
    public AbsWriteStyle<?> createDataColStyle(Supplier<AbsWriteStyle<?>> createWriteStyle) {
        if (this.dataColStyle == null) {
            this.dataColStyle = createWriteStyle.get();
            this.dataColStyle.init(getDataStyleConfig());
        }
        return this.dataColStyle;
    }

    public AbsWriteStyle<?> getDataColStyleCreate(Supplier<AbsWriteStyle<?>> createWriteStyle, long y) {
        AbsWriteStyle<?> dataColStyle = this.createDataColStyle(createWriteStyle);
        if (!isNewStyle()) {
            return dataColStyle;
        }
        AbsWriteStyle<?> newDataColStyle = dataColStyle.copy();
        this.createDataCelStyleMap.put(y, newDataColStyle);
        return newDataColStyle;
    }

    public AbsWriteStyle<?> copyDataColStyleCreate(Supplier<AbsWriteStyle<?>> createWriteStyle, long y) {
        AbsWriteStyle<?> dataColStyle = this.createDataColStyle(createWriteStyle);
        AbsWriteStyle<?> newDataColStyle = dataColStyle.copy();
        this.createDataCelStyleMap.put(y, newDataColStyle);
        return newDataColStyle;
    }

    public boolean isDraw() {
        return drawField != null;
    }

    public boolean isNotDraw() {
        return !this.isDraw();
    }
}
