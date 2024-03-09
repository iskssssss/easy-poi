package top.kongsheng.common.easy_poi.reader.core;

import top.kongsheng.common.easy_poi.reader.abs.AbsDataReader;

import java.io.Serializable;

/**
 * ReadResultCallback
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/5/19 10:02
 */
@FunctionalInterface
public interface ReadResultCallback<T> extends Serializable {

    /**
     * 处理
     * @param dataReader
     */
    void handle(AbsDataReader<T, ?> dataReader) throws InterruptedException;
}
