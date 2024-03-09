package top.kongsheng.common.easy_poi.value.convert;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.date.LocalDateTimeUtil;
import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.model.Cell;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

/**
 * WriteValueConvertDate
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/22 17:29
 */
public class WriteValueConvertDate extends WriteValueConvert {

    @Override
    public void write(Cell<?> cell, PoiModelField poiItem, Object value) {
        String cellValue = this.handlerDateData(poiItem, value);
        cell.writeString(poiItem, cellValue);
    }

    private String handlerDateData(PoiModelField poiItem, Object value) {
        PoiModelField.WriteConfig writeConfig = poiItem.writeConfig();
        String dateFormat = writeConfig.dateFormat();
        if (StrUtil.isBlank(dateFormat)) {
            dateFormat = "yyyy-MM-dd HH:mm:ss";
        }
        if (value instanceof Date) {
            return DateUtil.format((Date) value, dateFormat);
        } else if (value instanceof LocalDate) {
            return LocalDateTimeUtil.format((LocalDate) value, dateFormat);
        } else if (value instanceof LocalDateTime) {
            return DateUtil.format((LocalDateTime) value, dateFormat);
        } else {
            return null;
        }
    }
}
