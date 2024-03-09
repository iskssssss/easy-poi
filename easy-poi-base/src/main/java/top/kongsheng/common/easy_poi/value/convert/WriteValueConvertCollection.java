package top.kongsheng.common.easy_poi.value.convert;

import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.model.Cell;

import java.util.List;

/**
 * WriteValueConvertCollection
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/22 17:28
 */
public class WriteValueConvertCollection extends WriteValueConvert {

    @Override
    public void write(Cell<?> cell, PoiModelField poiItem, Object value) {
        List<?> collection = (List<?>) value;
        cell.writeList(poiItem, collection);
    }
}
