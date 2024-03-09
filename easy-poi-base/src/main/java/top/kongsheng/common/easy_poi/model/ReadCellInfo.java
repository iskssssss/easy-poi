package top.kongsheng.common.easy_poi.model;

import java.io.Closeable;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 读取数据信息
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/7 19:16
 */
public class ReadCellInfo implements Closeable  {
    private final Object value;
    private final Map<String, Object> paramsMap = new LinkedHashMap<>();

    public ReadCellInfo(Object value) {
        this.value = value;
    }

    public ReadCellInfo addParamsMap(String key, Object value) {
        this.paramsMap.put(key, value);
        return this;
    }

    public void setBackgroundColor(Object value) {
        this.paramsMap.put("BACKGROUND_COLOR", value);
    }

    public Object getBackgroundColor() {
        return this.paramsMap.get("BACKGROUND_COLOR");
    }

    public Object getValue() {
        return value;
    }

    public Map<String, Object> getParamsMap() {
        return paramsMap;
    }

    @Override
    public void close() {
        if (this.value instanceof Collection) {
            ((Collection<?>) this.value).clear();
        }
        this.paramsMap.clear();
    }
}
