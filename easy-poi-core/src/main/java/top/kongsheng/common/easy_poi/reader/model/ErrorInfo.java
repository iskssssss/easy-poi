package top.kongsheng.common.easy_poi.reader.model;

import cn.hutool.core.util.ObjectUtil;
import lombok.Data;

/**
 * 错误信息
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/4 14:30
 */
@Data
public abstract class ErrorInfo {
    private final String tableIndex;
    private final int x;
    private final int y;
    private final String title;
    private final Object value;
    private final String message;

    public ErrorInfo(String tableIndex, int x, int y, String title, Object value, String message) {
        this.tableIndex = tableIndex;
        this.x = x;
        this.y = y;
        this.title = title;
        this.value = value;
        this.message = message;
    }

    public String getVal() {
        if (ObjectUtil.isEmpty(this.value)) {
            return "(" + this.message + ")";
        }
        return this.value + "(" + this.message + ")";
    }

    public static final class ErrorInfoDefault extends ErrorInfo {

        public ErrorInfoDefault(String tableIndex, int xIndex, int yIndex, String title, Object value, String message) {
            super(tableIndex, xIndex, yIndex, title, value, message);
        }
    }

    public static ErrorInfoDefault defaultError(String tableIndex, int xIndex, int yIndex, String title, Object value, String message) {
        return new ErrorInfoDefault(tableIndex, xIndex, yIndex, title, value, message);
    }
}
