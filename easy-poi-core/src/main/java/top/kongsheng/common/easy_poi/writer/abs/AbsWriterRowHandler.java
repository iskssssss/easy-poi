package top.kongsheng.common.easy_poi.writer.abs;

import top.kongsheng.common.easy_poi.writer.model.RowHandleResult;

import java.io.Closeable;
import java.util.List;

/**
 * AbsWriterRowHandler
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/23 19:54
 */
public abstract class AbsWriterRowHandler<IN_TYPE, OUT_TYPE> implements Closeable {
    private final List<IN_TYPE> dataList;

    protected AbsWriterRowHandler(List<IN_TYPE> dataList) {
        this.dataList = dataList;
    }

    /**
     * 导出实体类类型
     *
     * @return 导出实体类类型
     */
    public abstract Class<OUT_TYPE> exportModelClass();


    /**
     * 转换实体类
     *
     * @param index  行数
     * @param inType 类型
     * @return 结果
     */
    protected abstract RowHandleResult<OUT_TYPE> to(long index, IN_TYPE inType);

    public List<IN_TYPE> getDataList() {
        return dataList;
    }

    @Override
    public void close() {
        dataList.clear();
    }
}
