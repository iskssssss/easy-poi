package top.kongsheng.common.easy_poi.enums;

import cn.hutool.core.util.ObjectUtil;

/**
 * CellStyleTypeEnum
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/10/27 14:46
 */
public enum CellStyleTypeEnum {
    /**
     * 大标题样式类型
     */
    MAIN_TITLE("MAIN_TITLE_", "大标题样式类型"),
    /**
     * 数据标题样式类型
     */
    DATA_TITLE("DATA_TITLE_", "数据标题样式类型"),
    /**
     * 数据样式类型
     */
    DATA("DATA_", "数据样式类型");
    private final String code;
    private final String text;

    CellStyleTypeEnum(String code, String text) {
        this.code = code;
        this.text = text;
    }

    public String inPlaceholder(Object placeholder) {
        if (ObjectUtil.isEmpty(placeholder)) {
            return this.code;
        }
        return this.code + placeholder.toString();
    }

    public String getCode() {
        return code;
    }

    public String getText() {
        return text;
    }
}
