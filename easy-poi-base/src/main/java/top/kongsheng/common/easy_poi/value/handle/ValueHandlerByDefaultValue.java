package top.kongsheng.common.easy_poi.value.handle;

import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.easy_poi.model.ReadCellInfo;

import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;

/**
 * ValueHandlerByDefaultValue
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/5 15:43
 */
public class ValueHandlerByDefaultValue implements ValueHandler<String> {
    private final Map<String, String> defaultValueMap = new ConcurrentHashMap<>(4);

    @Override
    public String to(String title, ReadCellInfo readCellInfo, Object... otherParams) {
        Object value = readCellInfo.getValue();
        String result = StrUtil.toString(value);
        if (value instanceof Collection) {
            result = ((List<?>) value).stream().map(StrUtil::toString).filter(item -> StrUtil.isNotBlank(item) && !"null".equals(item)).distinct().collect(Collectors.joining());
        }
        if (StrUtil.isBlank(result)) {
            return StrUtil.toString(defaultValueMap.get(title));
        }
        result = result.replaceAll(" ", "").replaceAll("\r", "").replaceAll("\n", "");
        defaultValueMap.put(title, result);
        return result;
    }

    @Override
    public void clear() {

    }
}
