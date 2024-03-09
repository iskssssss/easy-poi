package top.kongsheng.common.easy_poi.writer;

import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.cell.CellUtil;
import top.kongsheng.common.easy_poi.model.ExcelCell;
import top.kongsheng.common.easy_poi.utils.PaperSizeUtil;
import top.kongsheng.common.easy_poi.config.StyleConfig;
import top.kongsheng.common.easy_poi.enums.BorderStyle;
import top.kongsheng.common.easy_poi.handler.AbsWriteStyle;
import top.kongsheng.common.easy_poi.model.DrawFieldInfo;
import top.kongsheng.common.easy_poi.model.MergePosition;
import top.kongsheng.common.easy_poi.utils.CharUtil;
import top.kongsheng.common.easy_poi.writer.abs.AbsDataWriter;
import top.kongsheng.common.easy_poi.writer.abs.AbsWriterRowHandler;
import top.kongsheng.common.easy_poi.writer.model.RowHandleResult;
import top.kongsheng.common.easy_poi.writer.style.ExcelWriteStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Collection;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * 数据导出 Excel实现
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/11 17:44
 */
public class ExcelDataWriter<IN_TYPE, OUT_TYPE> extends AbsDataWriter<IN_TYPE, OUT_TYPE> {

    private static final float MARGIN_CONVERT_RATE = 1444.3329989969912047074020456556F;

    private final Workbook workbook;
    private Sheet sheet;

    public ExcelDataWriter(AbsWriterRowHandler<IN_TYPE, OUT_TYPE> rowHandler) throws IOException {
        super(rowHandler);
        super.setDataTitleRowPosition(poiModel.titleRowIndex());
        this.workbook = WorkbookFactory.create(true);
    }

    @Override
    protected void init() {
        String mainTitle = mainWriteConfig.getMainTitle();
        String sheetName = mainTitle.replaceAll("[\\\\/]", "_");
        this.sheet = this.workbook.createSheet(sheetName);
        boolean landscape = mainWriteConfig.isLandscape();
        this.sheet.setMargin(PageMargin.LEFT, (mainWriteConfig.getMarLeft() / MARGIN_CONVERT_RATE) / 2);
        this.sheet.setMargin(PageMargin.RIGHT, (mainWriteConfig.getMarRight() / MARGIN_CONVERT_RATE) / 2);
        this.sheet.setMargin(PageMargin.TOP, mainWriteConfig.getMarTop() / MARGIN_CONVERT_RATE);
        this.sheet.setMargin(PageMargin.BOTTOM, mainWriteConfig.getMarBottom() / MARGIN_CONVERT_RATE);
        this.sheet.setHorizontallyCenter(true);
        PrintSetup printSetup = this.sheet.getPrintSetup();
        PaperSizeUtil.PaperSizeEnum paperSizeEnum = PaperSizeUtil.find(mainWriteConfig.getDocumentWidth(), mainWriteConfig.getDocumentHeight(), landscape);
        if (paperSizeEnum != null) {
            printSetup.setPaperSize(paperSizeEnum.getSize());
        }
        if (landscape) {
            printSetup.setLandscape(true);
        }
        int titleIndex = super.dataTitleRowPosition - 1;
        if (titleIndex < 0) {
            return;
        }
        StyleConfig styleConfig = mainWriteConfig.getMainTitleStyle();
        // 绘制大标题
        ExcelWriteStyle mainTitleStyle = new ExcelWriteStyle(this.workbook);
        mainTitleStyle
                .setFont(styleConfig.getFont())
                .setFontBold(styleConfig.isFontBold())
                .setFontColor(styleConfig.getFontColor())
                .setFontSizePoints(styleConfig.getFontSizePoints())
                .setHorizontal(styleConfig.getHorizontal())
                .setVertical(styleConfig.getVertical())
                .setBorder(BorderStyle.NONE);
        this.setCellValue(titleIndex, 0, mainTitle, mainTitleStyle);
        this.setRowHeight(titleIndex, 800);
        int drawXDrawFieldInfoMapSize = this.drawXDrawFieldInfoMap.size();
        if (drawXDrawFieldInfoMapSize <= 1) {
            return;
        }
        CellRangeAddress region = new CellRangeAddress(titleIndex, titleIndex, 0, drawXDrawFieldInfoMapSize - 1);
        this.sheet.addMergedRegion(region);
    }

    @Override
    protected AbsWriteStyle<?> createWriteStyle() {
        return new ExcelWriteStyle(this.workbook);
    }

    @Override
    protected void setCellValue(int y, int x, Object value, AbsWriteStyle<?> writeStyle) {
        Row row = sheet.getRow(y) == null ? sheet.createRow(y) : sheet.getRow(y);
        Cell cell = row.getCell(x) == null ? row.createCell(x) : row.getCell(x);
        String valueStr = ObjectUtil.isEmpty(value) ? "" : value.toString();
        cell.setCellValue(valueStr);
        cell.setCellStyle((CellStyle) writeStyle.get());
    }

    @Override
    protected int writeData(int rowIndex, RowHandleResult<OUT_TYPE> rowHandleResult) {
        Collection<List<DrawFieldInfo>> drawOrderFieldInfoMapValues = this.drawXDrawFieldInfoMap.values();
        Row row = this.sheet.getRow(rowIndex) == null ? this.sheet.createRow(rowIndex) : this.sheet.getRow(rowIndex);
        ExcelCell wordCell = new ExcelCell();
        OUT_TYPE data = rowHandleResult.getData();
        for (List<DrawFieldInfo> drawFieldInfoList : drawOrderFieldInfoMapValues) {
            int drawX = drawFieldInfoList.iterator().next().getDrawX();
            Cell cell = row.getCell(drawX) == null ? row.createCell(drawX) : row.getCell(drawX);
            wordCell.setCell(cell);
            wordCell.setDrawFieldInfoList(drawFieldInfoList);
            for (DrawFieldInfo drawFieldInfo : drawFieldInfoList) {
                boolean nested = drawFieldInfo.isNested();
                if (nested) {
                    Object value;
                    try {
                        Field drawField = drawFieldInfo.getDrawField();
                        drawField.setAccessible(true);
                        value = drawField.get(data);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                        continue;
                    }
                    if (value instanceof Collection) {
                        Collection<?> collection = (Collection<?>) value;
                        List<DrawFieldInfo> childDrawFieldInfoList = drawFieldInfo.getChild();
                        ExcelCell childWordCell = new ExcelCell();
                        childWordCell.setCell(cell);
                        childWordCell.setDrawFieldInfoList(childDrawFieldInfoList);
                        for (Object obj : collection) {
                            for (DrawFieldInfo childDrawFieldInfo : childDrawFieldInfoList) {
                                childDrawFieldInfo.setCellRowMaxCharCount(drawFieldInfo.getCellRowMaxCharCount());
                                AbsWriteStyle<?> dataColStyle = childDrawFieldInfo.getDataColStyleCreate(this::createWriteStyle, rowIndex);
                                String fontColor = rowHandleResult.getFontColor();
                                if (StrUtil.isNotEmpty(fontColor)) {
                                    dataColStyle = childDrawFieldInfo.copyDataColStyleCreate(this::createWriteStyle, rowIndex);
                                    dataColStyle.setFontColor(fontColor);
                                }
                                childWordCell.initStyle(dataColStyle);
                                super.setCellValue(rowIndex, drawX, childWordCell, childDrawFieldInfo, obj);
                            }
                        }
                        continue;
                    }
                }
                AbsWriteStyle<?> dataColStyle = drawFieldInfo.getDataColStyleCreate(this::createWriteStyle, rowIndex);
                String fontColor = rowHandleResult.getFontColor();
                if (StrUtil.isNotEmpty(fontColor)) {
                    dataColStyle = drawFieldInfo.copyDataColStyleCreate(this::createWriteStyle, rowIndex);
                    dataColStyle.setFontColor(fontColor);
                }
                wordCell.initStyle(dataColStyle);
                if (drawFieldInfo.isCustom()) {
                    continue;
                }
                super.setCellValue(rowIndex, drawX, wordCell, drawFieldInfo, data);
            }
        }
        return 1;
    }

    @Override
    protected void merge(MergePosition mergePosition) {
        int startY = mergePosition.getStartY(), endY = mergePosition.getEndY();
        int startX = mergePosition.getStartX(), endX = mergePosition.getEndX();
        CellRangeAddress region = new CellRangeAddress(startY, endY, startX, endX);
        this.sheet.addMergedRegion(region);
    }

    @Override
    public void write(OutputStream outputStream) throws IOException {
        workbook.write(outputStream);
    }

    @Override
    protected void autoRowHeight(int y, boolean autoHeight, float defaultHeight, float heightRate) {
        Row row = sheet.getRow(y);
        float maxRowHeight = -1;
        if (autoHeight) {
            short minColIx = row.getFirstCellNum();
            short maxColIx = row.getLastCellNum();
            for (short colIx = minColIx; colIx < maxColIx; colIx++) {
                Cell cell = row.getCell(colIx);
                if (cell == null) {
                    continue;
                }
                CellStyle cellStyle = cell.getCellStyle();
                cellStyle.setWrapText(true);
                String text = StrUtil.toString(CellUtil.getCellValue(cell));
                List<DrawFieldInfo> drawFieldInfoList = this.drawXDrawFieldInfoMap.get(((int) colIx));
                final DrawFieldInfo drawFieldInfo = drawFieldInfoList.iterator().next();
                if (drawFieldInfo.isCustom()) {
                    continue;
                }
                AbsWriteStyle<?> dataColStyle = drawFieldInfo.createDataColStyle(this::createWriteStyle);
                float newRowHeight = this.textAutoHeight(colIx, text, heightRate, dataColStyle);
                maxRowHeight = Math.max(maxRowHeight, newRowHeight);
            }
        } else {
            maxRowHeight = defaultHeight;
        }
        row.setHeight((short) (maxRowHeight));
    }

    @Override
    public void close() {
        try {
            super.close();
            workbook.close();
        } catch (IOException ioException) {
            ioException.printStackTrace();
        }
    }

    @Override
    public int getTableWidth() {
        int documentWidthNotMargin = mainWriteConfig.getDocumentWidth();
        boolean landscape = mainWriteConfig.isLandscape();
        BigDecimal resultBigDecimal = BigDecimal.valueOf(documentWidthNotMargin * (landscape ? 2.125 : 2));
        int result = resultBigDecimal.setScale(0, RoundingMode.UP).intValue();
        return result;
    }

    @Override
    protected void setColumnWidth(int y, int x, float width, boolean gridCol) {
        if (width < 0) {
            return;
        }
        sheet.setColumnWidth(x, (int) width);
    }

    @Override
    protected void setRowHeight(int y, float height) {
        Row row = sheet.getRow(y) == null ? sheet.createRow(y) : sheet.getRow(y);
        row.setHeight((short) height);
    }

    protected float textAutoHeight(int x, String text, float engCountCopies, AbsWriteStyle<?> writeStyle) {
//        AtomicInteger newlineSum = new AtomicInteger();
//        long dataCharLength = CharUtil.dataCharLength(text, character -> {
//            if ('\n' == character) {
//                newlineSum.getAndIncrement();
//            }
//        });
//        float fontSizePoints = writeStyle.getFontSizePoints();
//        float fontSize = fontSizePoints * 20F;
//        float rowMaxCharNum = rowMaxCharNumMap.computeIfAbsent(x, key -> {
//            float widthInPixels = sheet.getColumnWidth(key);
//            float r = (widthInPixels - 1000) / fontSize;
//            return NumberUtil.round(r, 0, RoundingMode.UP).floatValue();
//        });
//        BigDecimal bigDecimal = NumberUtil.round(dataCharLength / rowMaxCharNum, 0, RoundingMode.UP);
//        float row = bigDecimal.floatValue();
//        float newRowHeight = ((row) * engCountCopies) * (fontSize + (65 + (fontSize / 4)));
//        return newRowHeight;
        AtomicInteger newlineSum = new AtomicInteger();
        long dataCharLength = CharUtil.dataCharLength(text, character -> {
            if ('\n' == character) {
                newlineSum.getAndIncrement();
            }
        });
        float cellRowMaxCharCount = this.getCellRowMaxCharCount(x);
//        cellRowMaxCharCount = cellRowMaxCharCount - ((dataCharLength / cellRowMaxCharCount) * 0.2F);
        float contentRow = 1F;
        if (dataCharLength >= cellRowMaxCharCount) {
            BigDecimal bigDecimal = BigDecimal.valueOf(dataCharLength / cellRowMaxCharCount);
            contentRow = bigDecimal.setScale(0, RoundingMode.UP).floatValue();
        }
        contentRow += (newlineSum.get() / 2F);
        float fontSizePt = writeStyle.getFontSizePoints() * 2;
        float rate = 0;
//        if (contentRow > 20) {
//            rate = 0.25F;
//        }
        float floatValue = BigDecimal.valueOf(contentRow * rate).setScale(0, RoundingMode.DOWN).floatValue();
        float contentHeight = (((contentRow - floatValue) * engCountCopies) * (fontSizePt * 9.5F)) * 2F;
        return contentHeight;
    }

    @Override
    protected float _getCellRowMaxCharCount(int x) {
        List<DrawFieldInfo> drawFieldInfoList = this.drawXDrawFieldInfoMap.get(x);
        final DrawFieldInfo drawFieldInfo = drawFieldInfoList.iterator().next();
        AbsWriteStyle<?> dataColStyle = drawFieldInfo.createDataColStyle(this::createWriteStyle);
        float fontSize = dataColStyle.getFontSizePoints();
        float widthInPixels = sheet.getColumnWidth(x) / 20F;
        return NumberUtil.round(widthInPixels / fontSize, 0, RoundingMode.UP).floatValue();
    }
}
