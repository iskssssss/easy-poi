package top.kongsheng.common.easy_poi.reader.model;

import cn.hutool.core.util.ObjectUtil;
import cn.hutool.core.util.ReflectUtil;
import top.kongsheng.common.easy_poi.reader.exception.ErrorInfoException;
import top.kongsheng.common.easy_poi.reader.handler.ImportInfo;
import top.kongsheng.common.easy_poi.value.handle.ValueHandler;
import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.model.ReadCellInfo;
import lombok.Data;

import java.lang.reflect.Field;

/**
 * 单元格信息
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/4 14:30
 */
@Data
public class CellInfo {
    /**
     * 导入信息
     */
    private final ImportInfo<?, ?> importInfo;
    /**
     * 字段注解
     */
    private PoiModelField poiModelField;
    /**
     * 字段信息
     */
    private final Field field;
    private final int xIndex;
    private final boolean required;
    private final ValueHandler valueHandler;
    private boolean write = false;

    public CellInfo(ImportInfo<?, ?> importInfo, Field field) {
        this.importInfo = importInfo;
        this.field = field;
        this.poiModelField = field.getAnnotation(PoiModelField.class);
        this.xIndex = poiModelField.x();
        PoiModelField.ReadConfig readConfig = this.poiModelField.readConfig();
        this.required = readConfig.required();
        this.valueHandler = importInfo.getTypeHandlerMap(readConfig.typeHandler());
    }

    public Object setValue(Object obj, String title, String tableIndex,
                         ReadCellInfo readCellInfo, int yIndex, boolean write,
                         Object... otherParams) throws ErrorInfoException {
        this.write = write;
        return this.setValue(obj, title, tableIndex, readCellInfo, yIndex, otherParams);
    }

    public Object setValue(Object obj, String title, String tableIndex,
                         ReadCellInfo readCellInfo, int yIndex,
                         Object... otherParams) throws ErrorInfoException {
        if (this.write) {
            return readCellInfo.getValue();
        }
        if (ObjectUtil.isEmpty(readCellInfo.getValue()) && this.required) {
            ErrorInfoException._throws(ErrorInfo.defaultError(tableIndex, this.xIndex, yIndex, title, readCellInfo.getValue(), "不可为空"));
        }
        this.valueHandler.clear();
        Object toValue = this.valueHandler.to(title, readCellInfo, otherParams);
        if (ObjectUtil.isEmpty(toValue) && this.required) {
            ErrorInfoException._throws(ErrorInfo.defaultError(tableIndex, this.xIndex, yIndex, title, toValue, "不可为空"));
        }
        ReflectUtil.setFieldValue(obj, this.field, toValue);
        write = true;
        return toValue;
    }
}
