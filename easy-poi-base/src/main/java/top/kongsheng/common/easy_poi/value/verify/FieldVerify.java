package top.kongsheng.common.easy_poi.value.verify;

import cn.hutool.core.util.ArrayUtil;

import java.io.Serializable;

/**
 * 导入数据校验
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/9/6 10:10
 */
@FunctionalInterface
public interface FieldVerify<FIELD_TYPE> extends Serializable {

    /**
     * 校验字段
     *
     * @param obj 值
     * @return 错误信息 返回空为正确
     */
    String verify(FIELD_TYPE obj);

    class FieldVerifyDefault implements FieldVerify<Object> {

        @Override
        public String verify(Object obj) {
            return null;
        }
    }

    class FieldVerifyNotEmpty implements FieldVerify<Object> {

        @Override
        public String verify(Object obj) {
            if (obj == null || "".equals(obj) || ArrayUtil.isEmpty(obj)) {
                return "不可为空";
            }
            return null;
        }
    }
}
