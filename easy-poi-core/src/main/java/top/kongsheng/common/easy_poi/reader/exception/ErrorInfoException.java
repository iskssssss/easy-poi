package top.kongsheng.common.easy_poi.reader.exception;

import top.kongsheng.common.easy_poi.reader.model.ErrorInfo;

/**
 * 数据异常
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/5/5 14:59
 */
public class ErrorInfoException extends RuntimeException {
    private final ErrorInfo errorInfo;

    public ErrorInfoException(ErrorInfo errorInfo) {
        this.errorInfo = errorInfo;
    }

    public ErrorInfo getErrorInfo() {
        return errorInfo;
    }

    public static void _throws(ErrorInfo errorInfo) throws ErrorInfoException {
        throw new ErrorInfoException(errorInfo);
    }
}
