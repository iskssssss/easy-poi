package top.kongsheng.common.easy_poi.value;

import cn.hutool.core.util.ReflectUtil;
import top.kongsheng.common.easy_poi.value.convert.*;
import top.kongsheng.common.easy_poi.value.convert.*;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * WriteValueConvertUtil
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/22 17:32
 */
public class WriteValueConvertUtil {

    public static Map<String, WriteValueConvert> WRITE_VALUE_CONVERT_MAP = new HashMap<>(8);

    static {
        WRITE_VALUE_CONVERT_MAP.put(WriteValueConvertDefault.class.getSimpleName(), new WriteValueConvertDefault());
        WRITE_VALUE_CONVERT_MAP.put(WriteValueConvertDate.class.getSimpleName(), new WriteValueConvertDate());
        WRITE_VALUE_CONVERT_MAP.put(WriteValueConvertCollection.class.getSimpleName(), new WriteValueConvertCollection());
        WRITE_VALUE_CONVERT_MAP.put(WriteValueConvertString.class.getSimpleName(), new WriteValueConvertString());
    }

    public static WriteValueConvert getWriteValueConvert(Class<? extends WriteValueConvert> writeValueConvertClass) {
        WriteValueConvert writeValueConvert = WRITE_VALUE_CONVERT_MAP.get(writeValueConvertClass.getSimpleName());
        if (writeValueConvert == null) {
            synchronized (WriteValueConvertUtil.class) {
                if (writeValueConvert == null) {
                    writeValueConvert = WriteValueConvertUtil.addWriteValueConvert(writeValueConvertClass);
                }
            }
        }
        return writeValueConvert;
    }

    public static WriteValueConvert addWriteValueConvert(Class<? extends WriteValueConvert> writeValueConvertClass) {
        WriteValueConvert writeValueConvert = ReflectUtil.newInstance(writeValueConvertClass);
        WRITE_VALUE_CONVERT_MAP.put(writeValueConvertClass.getSimpleName(), writeValueConvert);
        return writeValueConvert;
    }

    public static WriteValueConvert getWriteValueConvert(Object value) {
        if (value instanceof Date || value instanceof LocalDate || value instanceof LocalDateTime) {
            return WriteValueConvertUtil.getWriteValueConvert(WriteValueConvertDate.class);
        } else if (value instanceof Collection) {
            return WriteValueConvertUtil.getWriteValueConvert(WriteValueConvertCollection.class);
        } else {
            return WriteValueConvertUtil.getWriteValueConvert(WriteValueConvertString.class);
        }
    }
}
