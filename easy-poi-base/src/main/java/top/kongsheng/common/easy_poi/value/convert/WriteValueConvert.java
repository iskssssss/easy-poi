package top.kongsheng.common.easy_poi.value.convert;

import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.model.Cell;

/**
 * WriteValueConvert
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/22 17:27
 */
public abstract class WriteValueConvert {

    public abstract void write(Cell<?> cell, PoiModelField poiItem, Object value);
}
