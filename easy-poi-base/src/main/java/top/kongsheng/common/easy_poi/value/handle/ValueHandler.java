package top.kongsheng.common.easy_poi.value.handle;

import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.easy_poi.model.ReadCellInfo;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

/**
 * 类型转换器
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/4 12:54
 */
public interface ValueHandler<T> {

    /**
     * 转换
     *
     * @param title        对应标题
     * @param readCellInfo 值信息
     * @param otherParams  其它参数
     * @return 转换后的值
     */
    T to(String title, ReadCellInfo readCellInfo, Object... otherParams);

    void clear();

    class ValueHandlerDefault implements ValueHandler<Object> {
        @Override
        public Object to(String title, ReadCellInfo readCellInfo, Object... otherParams) {
            return readCellInfo.getValue();
        }

        @Override
        public void clear() {

        }
    }

    class ValueHandlerString implements ValueHandler<String> {

        @Override
        public String to(String title, ReadCellInfo readCellInfo, Object... otherParams) {
            Object value = readCellInfo.getValue();
            if (value == null) {
                return null;
            }
            if (value instanceof List) {
                List<?> valueList = (List<?>) value;
                String result = valueList.stream().map(StrUtil::toString).filter(StrUtil::isNotBlank).collect(Collectors.joining("\r\n"));
                return result;
            }
            return StrUtil.toString(value);
        }

        @Override
        public void clear() {

        }
    }

    class ValueHandlerArray implements ValueHandler<List<String>> {

        @Override
        public List<String> to(String title, ReadCellInfo readCellInfo, Object... otherParams) {
            Object value = readCellInfo.getValue();
            if (value instanceof List) {
                return ((List<String>) value).stream().filter(StrUtil::isNotBlank).collect(Collectors.toList());
            }
            if(value == null) {
                return Collections.emptyList();
            }
            String[] split = value.toString().split("\n");
            List<String> list = Arrays.stream(split).filter(StrUtil::isNotBlank).collect(Collectors.toList());
            return list;
        }

        @Override
        public void clear() {

        }
    }
}
