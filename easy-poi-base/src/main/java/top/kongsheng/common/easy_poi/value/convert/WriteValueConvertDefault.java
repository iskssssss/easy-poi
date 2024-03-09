package top.kongsheng.common.easy_poi.value.convert;

import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.model.Cell;
import top.kongsheng.common.easy_poi.value.WriteValueConvertUtil;

/**
 * WriteValueConvertDefault
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/22 17:44
 */
public class WriteValueConvertDefault extends WriteValueConvert {

    @Override
    public void write(Cell<?> cell, PoiModelField poiItem, Object value) {
        WriteValueConvertUtil.getWriteValueConvert(value).write(cell, poiItem, value);
    }
}
