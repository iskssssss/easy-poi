package top.kongsheng.common.easy_poi.writer.abs;

import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.easy_poi.utils.ClassType;
import top.kongsheng.common.easy_poi.utils.ClassTypeUtil;
import top.kongsheng.common.easy_poi.anno.PoiModel;
import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.config.StyleConfig;
import top.kongsheng.common.easy_poi.handler.AbsWriteStyle;
import top.kongsheng.common.easy_poi.model.Cell;
import top.kongsheng.common.easy_poi.model.DrawFieldInfo;
import top.kongsheng.common.easy_poi.model.MergePosition;
import top.kongsheng.common.easy_poi.model.PoiModelFieldWriteConfig;
import top.kongsheng.common.easy_poi.value.WriteValueConvertUtil;
import top.kongsheng.common.easy_poi.value.convert.WriteValueConvert;
import top.kongsheng.common.easy_poi.writer.config.DataWriterConfig;
import top.kongsheng.common.easy_poi.writer.config.MainWriteConfig;
import top.kongsheng.common.easy_poi.writer.model.RowHandleResult;

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;

/**
 * 数据导出器 抽象类
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/11 17:44
 */
public abstract class AbsDataWriter<IN_TYPE, OUT_TYPE> extends DataWriterConfig implements Closeable, Runnable {

    protected PoiModel poiModel;
    protected final AbsWriterRowHandler<IN_TYPE, OUT_TYPE> rowHandler;
    protected final Map<String, List<MergePosition>> mergePositionInfoMap = new HashMap<>();

    /**
     * 预设横坐标：绘制信息
     */
    protected final Map<Integer, List<DrawFieldInfo>> sourceXDrawFieldInfoMap = new LinkedHashMap<>();

    /**
     * 标题：绘制信息
     */
    protected final Map<String, DrawFieldInfo> titleDrawFieldInfoMap = new LinkedHashMap<>();

    /**
     * 绘制横坐标：绘制信息
     */
    protected final Map<Integer, List<DrawFieldInfo>> drawXDrawFieldInfoMap = new LinkedHashMap<>();

    /**
     * 自动宽度字段信息列表
     */
    protected final List<DrawFieldInfo> autoWidthDrawFieldInfoList = new LinkedList<>();

    /**
     * 绘制字段宽度比总和
     */
    protected float drawFieldWidthRateTotal = 0F;

    protected AbsDataWriter(AbsWriterRowHandler<IN_TYPE, OUT_TYPE> rowHandler) {
        this.rowHandler = rowHandler;
        this.poiModel = this.rowHandler.exportModelClass().getAnnotation(PoiModel.class);
    }

    /**
     * 初始化poi信息
     */
    private void initPoiInfo() {
        this.mainWriteConfig = MainWriteConfig.createByPoiModel(this.poiModel);
        Class<OUT_TYPE> exportModelClass = this.rowHandler.exportModelClass();
        // 添加数据标题行样式
        Field[] fields = exportModelClass.getDeclaredFields();
        StyleConfig dateStyleConfig = null, titleStyleConfig = null;
        int drawX = 0;
        for (Field field : fields) {
            PoiModelField poiModelField = field.getAnnotation(PoiModelField.class);
            if (poiModelField == null) {
                continue;
            }
            PoiModelField.WriteConfig writeConfig = poiModelField.writeConfig();
            int x = writeConfig.x();
            if (x == -1) {
                x = poiModelField.x();
                if (x == -1) {
                    continue;
                }
            }
            DrawFieldInfo fatherDrawFieldInfo = new DrawFieldInfo();
            List<DrawFieldInfo> drawFieldInfoList = this.sourceXDrawFieldInfoMap.computeIfAbsent(x, key -> new LinkedList<>());
            StyleConfig finalDataStyleConfig = dateStyleConfig = StyleConfig.createByPoiModelFieldWriteConfig(writeConfig);
            StyleConfig finalTitleStyleConfig = titleStyleConfig = finalDataStyleConfig.copy();
            fatherDrawFieldInfo.setSourceX(x);
            fatherDrawFieldInfo.setDrawSort(writeConfig.drawSort());
            fatherDrawFieldInfo.setWriteType(writeConfig.writeType());
            fatherDrawFieldInfo.setNewLine(writeConfig.newLine());
            fatherDrawFieldInfo.setTitle(poiModelField.value()[writeConfig.titleGetIndex()]);
            fatherDrawFieldInfo.setNewStyle(writeConfig.newStyle());
            fatherDrawFieldInfo.setTitleStyleConfig(finalTitleStyleConfig, true);
            fatherDrawFieldInfo.setDataStyleConfig(finalDataStyleConfig);
            fatherDrawFieldInfo.setDrawField(field);
            drawFieldInfoList.add(fatherDrawFieldInfo);
            fatherDrawFieldInfo.setWidthRate(writeConfig.widthRate());
            fatherDrawFieldInfo.setDrawPoiModelField(poiModelField);
            boolean autoWidth = writeConfig.autoWidth();
            float autoWidthRate = writeConfig.autoWidthRate();
            fatherDrawFieldInfo.setAutoWidth(autoWidth);
            fatherDrawFieldInfo.setAutoWidthRate(autoWidthRate);
            boolean nested = writeConfig.nested();
            fatherDrawFieldInfo.setNested(nested);
            if (nested) {
                List<DrawFieldInfo> child = new LinkedList<>();
                fatherDrawFieldInfo.setChild(child);
                ClassType classType = ClassTypeUtil.getSignatureValue(field, false, true);
                Class<?> type = field.getType();
                if (classType != null) {
                    type = classType.getTType();
                    ClassType vType = classType.getVType();
                    if (vType != null) {
                        type = vType.getTType();
                    }
                }
                Field[] declaredFields = type.getDeclaredFields();
                for (Field declaredField : declaredFields) {
                    PoiModelField modelField = declaredField.getAnnotation(PoiModelField.class);
                    if (modelField == null) {
                        continue;
                    }
                    PoiModelField.WriteConfig config = modelField.writeConfig();
                    DrawFieldInfo childDrawFieldInfo = new DrawFieldInfo();
                    StyleConfig childDataStyleConfig = StyleConfig.createByPoiModelFieldWriteConfig(config);
                    StyleConfig childTitleStyleConfig = childDataStyleConfig.copy();
                    childDrawFieldInfo.setSourceX(x);
                    childDrawFieldInfo.setDrawSort(config.drawSort());
                    childDrawFieldInfo.setWriteType(config.writeType());
                    childDrawFieldInfo.setNewLine(config.newLine());
                    childDrawFieldInfo.setNewStyle(config.newStyle());
                    childDrawFieldInfo.setTitleStyleConfig(childDataStyleConfig, true);
                    childDrawFieldInfo.setDataStyleConfig(childTitleStyleConfig);
                    childDrawFieldInfo.setDrawField(declaredField);
                    childDrawFieldInfo.setDrawPoiModelField(modelField);
                    child.add(childDrawFieldInfo);
                }
            }
            String fieldName = field.getName();
            boolean containsResult = this.systemFieldList.contains(fieldName);
            if (this.exclude && containsResult) {
                // 当处于排除模式时，{systemFieldList}存在当前字段则跳过。
                continue;
            } else if (!this.exclude && !containsResult) {
                // 当处于限制模式时，{systemFieldList}不存在当前字段则跳过。
                continue;
            }
            int tDrawX = drawFieldInfoList.size() > 1 ? drawFieldInfoList.iterator().next().getDrawX() : drawX++;
            fatherDrawFieldInfo.setDrawX(tDrawX);
            this.drawXDrawFieldInfoMap.computeIfAbsent(tDrawX, key -> new LinkedList<>()).add(fatherDrawFieldInfo);
            this.titleDrawFieldInfoMap.put(fieldName, fatherDrawFieldInfo);
        }
        this.sourceXDrawFieldInfoMap.values().forEach(this::sort);
        // 自定义字段
        for (String customizeField : customizeFieldList) {
            DrawFieldInfo drawFieldInfo = new DrawFieldInfo();
            drawFieldInfo.setTitle(customizeField);
            drawFieldInfo.setCustom(true);
            drawFieldInfo.setWidthRate(0.02F);
            drawFieldInfo.setAutoWidth(true);
            drawFieldInfo.setTitleStyleConfig(titleStyleConfig.copy(), false);
            drawFieldInfo.setDataStyleConfig(dateStyleConfig.copy());
            drawFieldInfo.setDrawX(drawX++);
            final List<DrawFieldInfo> drawFieldInfoList = this.drawXDrawFieldInfoMap.computeIfAbsent(drawFieldInfo.getDrawX(), key -> new LinkedList<>());
            drawFieldInfoList.add(drawFieldInfo);
            this.titleDrawFieldInfoMap.put(customizeField, drawFieldInfo);
        }
        this.drawXDrawFieldInfoMap.values().forEach(this::sort);
        // 更新配置信息
        this.mainWriteConfig.update(this.customMainWriteConfig);
        Optional.ofNullable(this.customDrawFieldInfoList).ifPresent(list -> {
            for (PoiModelFieldWriteConfig poiModelFieldWriteConfig : list) {
                String drawFieldKey = poiModelFieldWriteConfig.getDrawFieldKey();
                if ("@ALL_UPDATE".equals(drawFieldKey)) {
                    Collection<DrawFieldInfo> titleDrawFieldInfoMapValues = this.titleDrawFieldInfoMap.values();
                    for (DrawFieldInfo drawFieldInfo : titleDrawFieldInfoMapValues) {
                        drawFieldInfo.update(poiModelFieldWriteConfig);
                    }
                    break;
                }
            }
            for (PoiModelFieldWriteConfig poiModelFieldWriteConfig : list) {
                String drawFieldKey = poiModelFieldWriteConfig.getDrawFieldKey();
                if ("@ALL_UPDATE".equals(drawFieldKey)) {
                    continue;
                }
                Optional.ofNullable(this.titleDrawFieldInfoMap.get(drawFieldKey))
                        .ifPresent(drawFieldInfo -> drawFieldInfo.update(poiModelFieldWriteConfig));
            }
        });
    }

    private void sort(List<DrawFieldInfo> drawFieldInfoList) {
        drawFieldInfoList.sort((o1, o2) -> {
            int drawSort1 = o1.getDrawSort();
            int drawSort2 = o2.getDrawSort();
            if (drawSort1 == drawSort2) {
                return 0;
            } else if (drawSort1 <= drawSort2) {
                return -1;
            }
            return 1;
        });
    }

    @Override
    public void run() {
        // 初始化poi信息
        this.initPoiInfo();
        // 自定义初始化
        this.init();
        // 绘制标题
        Collection<List<DrawFieldInfo>> drawXDrawFieldInfoMapValues = this.drawXDrawFieldInfoMap.values();
        for (List<DrawFieldInfo> drawFieldInfoList : drawXDrawFieldInfoMapValues) {
            final DrawFieldInfo drawFieldInfo = drawFieldInfoList.iterator().next();
            this.drawTitle(drawFieldInfo);
        }
        if (!this.drawXDrawFieldInfoMap.isEmpty()) {
            // 自动宽度
            Float diff = 1F - this.drawFieldWidthRateTotal;
            if (!diff.equals(0.0F)) {
                int documentWidth = this.getTableWidth();
                float wi = documentWidth * diff;
                List<DrawFieldInfo> tList = new LinkedList<>();
                for (DrawFieldInfo drawFieldInfo : autoWidthDrawFieldInfoList) {
                    Float autoWidthRate = drawFieldInfo.getAutoWidthRate();
                    if (autoWidthRate == null || autoWidthRate == -1) {
                        tList.add(drawFieldInfo);
                        continue;
                    }
                    float width = drawFieldInfo.getWidth();
                    float inWidth = wi * autoWidthRate;
                    float newWidth = width + inWidth;
                    drawFieldInfo.setWidth(newWidth);
                    wi -= inWidth;
                }
                if (!tList.isEmpty()) {
                    float widthAvg = wi / tList.size();
                    for (DrawFieldInfo drawFieldInfo : tList) {
                        float width = drawFieldInfo.getWidth();
                        float newWidth = width + widthAvg;
                        drawFieldInfo.setWidth(newWidth);
                    }
                }
            }
            Collection<List<DrawFieldInfo>> drawOrderFieldInfoMapValues = this.drawXDrawFieldInfoMap.values();
            for (List<DrawFieldInfo> drawFieldInfoList : drawOrderFieldInfoMapValues) {
                final DrawFieldInfo drawFieldInfo = drawFieldInfoList.iterator().next();
                this.setColumnWidth(super.dataTitleRowPosition, drawFieldInfo.getDrawX(), drawFieldInfo.getWidth(), true);
            }
            this.autoRowHeight(super.dataTitleRowPosition, true, -1, 1.5F);
            // 绘制数据
            List<IN_TYPE> dataList = rowHandler.getDataList();
            int size = dataList.size();
            for (int dataIndex = 0, dataRowIndex = 1, rowIndex = super.dataTitleRowPosition + 1; dataIndex < size; dataIndex++, dataRowIndex++) {
                RowHandleResult<OUT_TYPE> rowHandleResult = rowHandler.to(dataRowIndex, dataList.get(dataIndex));
                int increment = this.writeData(rowIndex, rowHandleResult);
                int rowHeight = this.poiModel.rowHeight();
                boolean autoHeight = rowHeight == -1;
                this.autoRowHeight(rowIndex, autoHeight, rowHeight * 36, 1F);
                rowIndex = rowIndex + increment;
            }
            this.merge(mergePositionInfoMap.values());
        }
        // 执行后置处理
        this.runEnd();
    }

    protected void runEnd() {

    }

    /**
     * 初始化
     */
    protected abstract void init();

    /**
     * 写入信息
     *
     * @param y          纵坐标
     * @param x          横坐标
     * @param value      值
     * @param writeStyle 单元格样式
     */
    protected abstract void setCellValue(int y, int x, Object value, AbsWriteStyle<?> writeStyle);

    /**
     * 写数据
     *
     * @param rowIndex        纵坐标
     * @param rowHandleResult 行数据
     * @return 递增值
     */
    protected abstract int writeData(int rowIndex, RowHandleResult<OUT_TYPE> rowHandleResult);

    /**
     * 合并单元格
     *
     * @param mergePositionList 待合并单元格信息
     */
    protected void merge(Collection<List<MergePosition>> mergePositionList) {
        for (List<MergePosition> mergePositions : mergePositionList) {
            if (mergePositions.isEmpty()) {
                continue;
            }
            for (MergePosition mergePosition : mergePositions) {
                int startY = mergePosition.getStartY(), endY = mergePosition.getEndY();
                if (startY == -1 || endY == -1) {
                    continue;
                }
                this.merge(mergePosition);
            }
        }
    }

    /**
     * 合并单元格
     *
     * @param mergePosition 待合并单元格信息
     */
    protected abstract void merge(MergePosition mergePosition);

    /**
     * 输出文件信息
     *
     * @param outputStream 输出流
     * @throws IOException 流异常
     */
    protected abstract void write(OutputStream outputStream) throws IOException;

    protected abstract AbsWriteStyle<?> createWriteStyle();

    @Override
    public void close() throws IOException {
        this.write(super.outputStream);
        this.rowHandler.close();
        this.mergePositionInfoMap.clear();
        this.customizeFieldList.clear();
        this.systemFieldList.clear();
        this.sourceXDrawFieldInfoMap.clear();
        this.titleDrawFieldInfoMap.clear();
        this.drawXDrawFieldInfoMap.clear();
        this.rowMaxCharNumMap.clear();
    }

    /**
     * 写标题
     *
     * @param drawFieldInfo 绘制信息
     * @return 最大高度
     */
    private void drawTitle(DrawFieldInfo drawFieldInfo) {
        String title = drawFieldInfo.getTitle();
        float width = this.getTableWidth() * drawFieldInfo.getWidthRate();
        int drawX = drawFieldInfo.getDrawX();
        drawFieldInfo.setWidth(width);
        AbsWriteStyle<?> titleColStyle = drawFieldInfo.createTitleColStyle(this::createWriteStyle);
        this.setCellValue(super.dataTitleRowPosition, drawX, title, titleColStyle);
        Float autoWidthRate = drawFieldInfo.getAutoWidthRate();
        if (drawFieldInfo.isAutoWidth() && (autoWidthRate == null || autoWidthRate != 0F)) {
            this.autoWidthDrawFieldInfoList.add(drawFieldInfo);
        }
        this.drawFieldWidthRateTotal += drawFieldInfo.getWidthRate();
    }

    /**
     * 表格宽度
     *
     * @return 表格宽度
     */
    public abstract int getTableWidth();

    /**
     * 处理并写单元格数据信息
     *
     * @param y               纵坐标
     * @param x               横坐标
     * @param cell            单元格信息
     * @param data 行数据
     */
    protected void setCellValue(int y, int x, Cell<?> cell, DrawFieldInfo drawFieldInfo, Object data) {
        this.getCellRowMaxCharCount(x);
        Field drawField = drawFieldInfo.getDrawField();
        PoiModelField drawPoiModelField = drawFieldInfo.getDrawPoiModelField();
        drawField.setAccessible(true);
        PoiModelField.WriteConfig writeConfig = drawPoiModelField.writeConfig();
        try {
            Object value = drawField.get(data);
            if (writeConfig.merge()) {
                boolean isContinue = this.mergeHandler(y, x, data, value, writeConfig);
                if (isContinue) {
                    return;
                }
            }
            Class<? extends WriteValueConvert> writeValueConvertClass = writeConfig.writeValueConvert();
            WriteValueConvert writeValueConvert = WriteValueConvertUtil.getWriteValueConvert(writeValueConvertClass);
            if (writeValueConvert == null) {
                writeValueConvert = WriteValueConvertUtil.getWriteValueConvert(value);
            }
            writeValueConvert.write(cell, drawPoiModelField, value);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    }

    /**
     * 合并单元格处理
     */
    private boolean mergeHandler(int y, int x, Object obj, Object cellValue, PoiModelField.WriteConfig writeConfig) {
        String mergeParentFieldValue = this.getMergeParentFieldValue(writeConfig, obj);
        String valueStr = cellValue == null ? "" : StrUtil.toString(cellValue);
        String mergeKeyStartText = x + mergeParentFieldValue;
        String currentKey = mergeKeyStartText + valueStr;
        List<MergePosition> mergePositionList = this.mergePositionInfoMap.computeIfAbsent(currentKey, item -> new ArrayList<>());
        MergePosition mergePosition = null;
        if (mergePositionList.isEmpty()) {
            for (Map.Entry<String, List<MergePosition>> entry : this.mergePositionInfoMap.entrySet()) {
                String key = entry.getKey();
                if (!key.startsWith(mergeKeyStartText)) {
                    continue;
                }
                List<MergePosition> value = entry.getValue();
                MergePosition temp = new MergePosition();
                if (Objects.equals(key, currentKey)) {
                    temp.setMergeText(valueStr);
                    mergePosition = temp;
                } else {
                    temp.setMergeText(value.get(0).getMergeText());
                }
                value.add(temp);
            }
        } else {
            int size = mergePositionList.size();
            mergePosition = mergePositionList.get(size - 1);
        }
        if (mergePosition == null) {
            throw new RuntimeException("合并单元格处理失败");
        }
        boolean isContinue = mergePosition.getStartY() != -1;
        mergePosition.setX(x);
        mergePosition.setAutoY(y);
        if (isContinue && !Objects.equals(valueStr, mergePosition.getMergeText())) {
            mergePosition = new MergePosition();
            mergePositionList.add(mergePosition);
            mergePosition.setMergeText(valueStr);
        }
        return isContinue;
    }

    /**
     * @param writeConfig
     * @param obj
     * @return
     */
    private String getMergeParentFieldValue(PoiModelField.WriteConfig writeConfig, Object obj) {
        StringBuilder mergeParentFieldValue = new StringBuilder();
        String[] mergeParentFieldNames = writeConfig.mergeParentFieldNames();
        if (mergeParentFieldNames.length <= 0) {
            return "";
        }
        for (String mergeParentFieldName : mergeParentFieldNames) {
            DrawFieldInfo drawFieldInfo = this.titleDrawFieldInfoMap.get(mergeParentFieldName);
            if (drawFieldInfo == null) {
                continue;
            }
            Field drawField = drawFieldInfo.getDrawField();
            try {
                Object value = drawField.get(obj);
                String valueStr = value == null ? "" : StrUtil.toString(value);
                mergeParentFieldValue.append(valueStr);
            } catch (IllegalAccessException ignored) {
            }
            break;
        }
        return mergeParentFieldValue.toString();
    }

    /**
     * 设置列宽
     *
     * @param y       行号
     * @param x       列号
     * @param width   宽度
     * @param gridCol 是否设置网格宽度
     */
    protected abstract void setColumnWidth(int y, int x, float width, boolean gridCol);

    /**
     * 设置行高
     *
     * @param y      行号
     * @param height 高度
     */
    protected abstract void setRowHeight(int y, float height);

    /**
     * 自动高度
     *
     * @param y             行号
     * @param autoHeight    是否自动高度
     * @param defaultHeight 默认行高
     * @param heightRate    高度比
     */
    protected void autoRowHeight(int y, boolean autoHeight, float defaultHeight, float heightRate) {
    }

    /**
     * 横坐标 ：可写入字符数量
     */
    private final Map<Integer, Float> rowMaxCharNumMap = new HashMap<>();

    protected float getCellRowMaxCharCount(int x) {
        return this.rowMaxCharNumMap.computeIfAbsent(x, key -> {
            float cellRowMaxCharCount = this._getCellRowMaxCharCount(key);
            List<DrawFieldInfo> drawFieldInfoList = this.drawXDrawFieldInfoMap.get(key);
            final DrawFieldInfo drawFieldInfo = drawFieldInfoList.iterator().next();
            drawFieldInfo.setCellRowMaxCharCount(cellRowMaxCharCount);
            return cellRowMaxCharCount;
        });
    }

    protected abstract float _getCellRowMaxCharCount(int x);
}
