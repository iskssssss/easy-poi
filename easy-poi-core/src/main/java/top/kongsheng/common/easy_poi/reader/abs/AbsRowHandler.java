package top.kongsheng.common.easy_poi.reader.abs;

import cn.hutool.core.util.ReflectUtil;
import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.easy_poi.reader.DataReadUtil;
import top.kongsheng.common.easy_poi.reader.exception.ErrorInfoException;
import top.kongsheng.common.easy_poi.reader.handler.ImportInfo;
import top.kongsheng.common.easy_poi.reader.model.CellInfo;
import top.kongsheng.common.easy_poi.reader.model.ErrorInfo;
import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.model.ReadCellInfo;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * 抽象行处理器
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/4 14:31
 */
public abstract class AbsRowHandler<T> {
    private final AtomicBoolean init = new AtomicBoolean(false);
    private ImportInfo<T, ?> importInfo;
    private Map<Integer, String> titleIndexMap;
    private Map<String, List<CellInfo>> titleCellInfoMap;
    private Boolean verify = false;

    /**
     * 初始化导入信息
     */
    public final void init(ImportInfo<T, ?> importInfo) {
        if (init.get()) {
            return;
        }
        init.set(true);
        this.importInfo = importInfo;
        List<String> modelTitleInfoList = importInfo.getModelTitleInfoList();
        this.titleCellInfoMap = importInfo.getTitleCellInfoMap();
        this.titleIndexMap = importInfo.getTitleIndexMap();
        Class<T> importModelClass = this.importModelClass();
        Field[] fields = importModelClass.getDeclaredFields();
        for (Field field : fields) {
            PoiModelField poiModelField = field.getAnnotation(PoiModelField.class);
            if (poiModelField == null || poiModelField.x() == -1) {
                continue;
            }
            String[] titles = poiModelField.value();
            for (String title : titles) {
                modelTitleInfoList.add(title);
                List<CellInfo> cellInfoList = this.titleCellInfoMap.computeIfAbsent(title, item -> new ArrayList<>());
                cellInfoList.add(new CellInfo(importInfo, field));
            }
        }
    }

    /**
     * 转换的实体类类型
     *
     * @return 实体类类型
     */
    protected abstract Class<T> importModelClass();

    public final T _to(String tableIndex, int rowIndex, List<ReadCellInfo> dataList, Object... otherParams) {
        if (!init.get()) {
            throw new IllegalArgumentException("未初始化");
        }
        return this.to(tableIndex, rowIndex, dataList, otherParams);
    }

    /**
     * 处理行数据
     *
     * @param tableIndex  表格编号
     * @param dataList    行数据
     * @param otherParams 其它参数
     * @return 处理结果
     */
    protected T to(String tableIndex, int rowIndex, List<ReadCellInfo> dataList, Object... otherParams) {
        if (DataReadUtil.isEmpty(dataList, this.titleIndexMap.size())) {
            return null;
        }
        T t = ReflectUtil.newInstance(this.importModelClass());
        for (int i = 0; i < dataList.size(); i++) {
            ReadCellInfo readCellInfo = dataList.get(i);
            String title = this.titleIndexMap.get(i);
            List<CellInfo> cellInfoList = this.titleCellInfoMap.get(title);
            for (CellInfo cellInfo : cellInfoList) {
                ErrorInfo errorInfo = null;
                Object value = null;
                try {
                    value = cellInfo.setValue(t, title, tableIndex, readCellInfo, rowIndex, false, otherParams);
                } catch (ErrorInfoException errorInfoException) {
                    errorInfo = errorInfoException.getErrorInfo();
                }
                if (!this.verify) {
                    continue;
                }
                if (errorInfo == null) {
                    String verify = this.verify(cellInfo.getPoiModelField(), value);
                    if (StrUtil.isNotEmpty(verify)) {
                        int xIndex = cellInfo.getXIndex();
                        errorInfo = new ErrorInfo.ErrorInfoDefault(tableIndex, xIndex, rowIndex, title, readCellInfo.getValue(), verify);
                    }
                }
                if (errorInfo == null) {
                    continue;
                }
                this.importInfo.getReadResult().addErrorInfoList(errorInfo);
            }
        }
        return t;
    }

    /**
     * 校验数据
     *
     * @param poiModelField 标题
     * @param value         数据
     */
    protected String verify(PoiModelField poiModelField, Object value) {
        return null;
    }

    /**
     * 全部读取完后的处理
     *
     * @param result 结果
     */
    public void afterHandler(List<T> result) {
    }

    public Boolean getVerify() {
        return this.verify;
    }

    public void setVerify(Boolean verify) {
        if (verify == null) {
            this.verify = false;
            return;
        }
        this.verify = verify;
    }
}
