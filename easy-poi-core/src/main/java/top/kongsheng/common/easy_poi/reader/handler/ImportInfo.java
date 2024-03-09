package top.kongsheng.common.easy_poi.reader.handler;

import cn.hutool.core.util.ReflectUtil;
import top.kongsheng.common.easy_poi.reader.ImportFileManager;
import top.kongsheng.common.easy_poi.reader.model.CellInfo;
import top.kongsheng.common.easy_poi.reader.model.ReadResult;
import top.kongsheng.common.easy_poi.value.handle.ValueHandler;

import java.io.Closeable;
import java.io.IOException;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 导入信息
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/7 12:15
 */
public class ImportInfo<T, MERGED_TYPE> implements Closeable, Serializable {
    private static final long serialVersionUID = 42L;
    /**
     * 标题
     */
    protected String topic;
    /**
     * 数据处理器
     */
    protected final Map<String, ValueHandler> typeHandlerMap = new ConcurrentHashMap<>(16);
    /**
     * 标题 ： 下标
     */
    protected final Map<Integer, String> titleIndexMap = new ConcurrentHashMap<>();
    /**
     * 标题 ： 单元格信息
     */
    protected final Map<String, List<CellInfo>> titleCellInfoMap = new ConcurrentHashMap<>();
    /**
     * 实体类标题信息
     */
    protected final List<String> modelTitleInfoList = new ArrayList<>();
    /**
     * 导入结果
     */
    protected final ReadResult<T> readResult;
    /**
     * 合并信息
     */
    private List<MERGED_TYPE> mergedRegions;

    public ImportInfo() {
        this.readResult = new ReadResult<>();
    }

    public ValueHandler getTypeHandlerMap(Class<? extends ValueHandler> typeHandlerClass) {
        String name = typeHandlerClass.getName();
        ValueHandler valueHandler = typeHandlerMap.computeIfAbsent(name, key -> ReflectUtil.newInstance(typeHandlerClass));
        return valueHandler;
    }

    public Map<String, ValueHandler> getTypeHandlerMap() {
        return typeHandlerMap;
    }

    public Map<Integer, String> getTitleIndexMap() {
        return titleIndexMap;
    }

    public Map<String, List<CellInfo>> getTitleCellInfoMap() {
        return titleCellInfoMap;
    }

    public List<String> getModelTitleInfoList() {
        return modelTitleInfoList;
    }

    public String getTopic() {
        return topic;
    }

    public void setTopic(String topic) {
        this.topic = topic;
    }

    public List<MERGED_TYPE> getMergedRegions() {
        return mergedRegions;
    }

    public void setMergedRegions(List<MERGED_TYPE> mergedRegions) {
        this.mergedRegions = mergedRegions;
    }

    public ReadResult<T> getReadResult() {
        return readResult;
    }

    @Override
    public void close() throws IOException {
        boolean error = this.readResult.isError();
        this.titleIndexMap.clear();
        this.titleCellInfoMap.values().parallelStream().forEach(List::clear);
        this.titleCellInfoMap.clear();
        this.modelTitleInfoList.clear();
        this.typeHandlerMap.values().parallelStream().forEach(ValueHandler::clear);
        this.typeHandlerMap.clear();
        this.mergedRegions.clear();
        if (error) {
            return;
        }
        String fileId = this.readResult.getFileId();
        ImportFileManager.del(fileId);
    }
}
