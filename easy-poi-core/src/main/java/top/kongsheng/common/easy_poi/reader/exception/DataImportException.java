package top.kongsheng.common.easy_poi.reader.exception;

/**
 * 数据导入异常类
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/7 17:46
 */
public class DataImportException extends RuntimeException {
    public DataImportException(String message) {
        super(message);
    }

    public DataImportException(Throwable cause) {
        super(cause);
    }

    public DataImportException(String message, Throwable cause) {
        super(message, cause);
    }
}
