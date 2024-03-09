package top.kongsheng.common.easy_poi.value.convert;

import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.model.Cell;

/**
 * WriteValueConvertString
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/22 17:28
 */
public class WriteValueConvertString extends WriteValueConvert {

    @Override
    public void write(Cell<?> cell, PoiModelField poiItem, Object value) {
        String cellValue = value == null ? "" : StrUtil.toString(value);
        cell.writeString(poiItem, cellValue);
    }
}
