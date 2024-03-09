package top.kongsheng.common.easy_poi.writer.config;

import top.kongsheng.common.easy_poi.model.PoiModelFieldWriteConfig;

import java.io.OutputStream;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * 数据导出配置信息
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/10/28 17:09
 */
public class DataWriterConfig implements Serializable {
    static final long serialVersionUID = 42L;

    /**
     * 数据标题开始位置
     */
    protected int dataTitleRowPosition;

    /**
     * 自定义字段名称
     */
    protected final List<String> customizeFieldList = new ArrayList<>();

    /**
     * 此字段表示'systemFieldList'中的字段是排除字段还是限制字段 true:排除字段 false:限制字段
     */
    protected boolean exclude = false;

    /**
     * 需要导出的字段
     */
    protected final List<String> systemFieldList = new ArrayList<>();

    protected OutputStream outputStream;

    protected MainWriteConfig mainWriteConfig;

    protected MainWriteConfig customMainWriteConfig;

    protected Collection<PoiModelFieldWriteConfig> customDrawFieldInfoList;

    /**
     * 加载自定义配置
     * @param customDrawFieldInfoList 自定义配置
     */
    public void loadCustomConfig(Collection<PoiModelFieldWriteConfig> customDrawFieldInfoList) {
        this.customDrawFieldInfoList = customDrawFieldInfoList;
    }

    public void loadCustomMainWriteConfig(MainWriteConfig customMainWriteConfig) {
        this.customMainWriteConfig = customMainWriteConfig;
    }

    public DataWriterConfig setExclude(boolean exclude) {
        this.exclude = exclude;
        return this;
    }

    public DataWriterConfig addCustomizeFieldList(List<String> customizeFieldList) {
        if (customizeFieldList == null) {
            return this;
        }
        this.customizeFieldList.addAll(customizeFieldList);
        return this;
    }

    public DataWriterConfig addSystemFieldList(List<String> systemFieldList) {
        if (systemFieldList == null) {
            return this;
        }
        this.systemFieldList.addAll(systemFieldList);
        return this;
    }

    public DataWriterConfig setOutputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
        return this;
    }

    public DataWriterConfig setDataTitleRowPosition(int dataTitleRowPosition) {
        this.dataTitleRowPosition = dataTitleRowPosition;
        return this;
    }
}
