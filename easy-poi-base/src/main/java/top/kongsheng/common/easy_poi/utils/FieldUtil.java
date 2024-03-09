package top.kongsheng.common.easy_poi.utils;

import java.util.function.Consumer;

/**
 * 字段工具类
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/10/13 17:03
 */
public class FieldUtil {


    /**
     * 更新字段的值
     *
     * @param value    新值
     * @param consumer 设置方法
     * @param <P>      类型
     */
    public static <P> void updateNotNull(P value, Consumer<P> consumer) {
        if (value == null) {
            return;
        }
        consumer.accept(value);
    }
}
