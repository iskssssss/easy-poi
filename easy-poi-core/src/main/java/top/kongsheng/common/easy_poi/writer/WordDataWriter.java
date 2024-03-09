package top.kongsheng.common.easy_poi.writer;

import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.easy_poi.model.WordCell;
import top.kongsheng.common.easy_poi.utils.WordStyleUtil;
import top.kongsheng.common.easy_poi.config.StyleConfig;
import top.kongsheng.common.easy_poi.enums.BorderStyle;
import top.kongsheng.common.easy_poi.handler.AbsWriteStyle;
import top.kongsheng.common.easy_poi.model.DrawFieldInfo;
import top.kongsheng.common.easy_poi.model.MergePosition;
import top.kongsheng.common.easy_poi.utils.CharUtil;
import top.kongsheng.common.easy_poi.utils.SizeConvertUtil;
import top.kongsheng.common.easy_poi.writer.abs.AbsDataWriter;
import top.kongsheng.common.easy_poi.writer.abs.AbsWriterRowHandler;
import top.kongsheng.common.easy_poi.writer.model.RowHandleResult;
import top.kongsheng.common.easy_poi.writer.style.WordWriteStyle;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.math.RoundingMode;
import java.util.Collection;
import java.util.List;
import java.util.Locale;

/**
 * 数据导出 Word实现
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/20 14:58
 */
public class WordDataWriter<IN_TYPE, OUT_TYPE> extends AbsDataWriter<IN_TYPE, OUT_TYPE> {

    private final XWPFDocument document;
    private XWPFTable table;

    public WordDataWriter(AbsWriterRowHandler<IN_TYPE, OUT_TYPE> rowHandler) {
        super(rowHandler);
        super.setDataTitleRowPosition(0);
        this.document = new XWPFDocument();
    }

    @Override
    protected void init() {
        CTSectPr sectPr = this.document.getDocument().getBody().addNewSectPr();
        boolean landscape = this.mainWriteConfig.isLandscape();
        if (landscape) {
            sectPr.addNewPgSz().setOrient(STPageOrientation.LANDSCAPE);
        }
        // 设置页面大小
        Integer documentWidth = this.mainWriteConfig.getDocumentWidth();
        Integer documentHeight = this.mainWriteConfig.getDocumentHeight();
        CTPageSz pgsz = sectPr.addNewPgSz();
        pgsz.setW(BigInteger.valueOf(documentWidth));
        pgsz.setH(BigInteger.valueOf(documentHeight));
        // 设置页面边距
        CTPageMar pageMar = sectPr.addNewPgMar();
        pageMar.setLeft(BigInteger.valueOf(this.mainWriteConfig.getMarLeft()));
        pageMar.setRight(BigInteger.valueOf(this.mainWriteConfig.getMarRight()));
        pageMar.setTop(BigInteger.valueOf(this.mainWriteConfig.getMarTop()));
        pageMar.setBottom(BigInteger.valueOf(this.mainWriteConfig.getMarBottom()));
        // 绘制大标题
        String mainTitle = this.mainWriteConfig.getMainTitle();
        XWPFParagraph paragraph = this.document.createParagraph();
        XWPFRun run = paragraph.createRun();
        StyleConfig mainTitleStyle = this.mainWriteConfig.getMainTitleStyle();
        run.setText(mainTitle);
        run.setTextPosition(mainTitleStyle.getTextPosition());
        run.setFontSize(mainTitleStyle.getFontSizePoints());
        run.setBold(mainTitleStyle.isFontBold());
        WordStyleUtil.setFont(run, mainTitleStyle.getFont());
        String vertical = mainTitleStyle.getVertical();
        ParagraphAlignment alignment = ParagraphAlignment.CENTER;
        if (StrUtil.isNotEmpty(vertical)) {
            alignment = ParagraphAlignment.valueOf(vertical.toUpperCase(Locale.ROOT));
        }
        paragraph.setAlignment(alignment);
        // 创建表格
        this.table = this.document.createTable();
        this.table.setTableAlignment(TableRowAlign.CENTER);
        int tableWidth = this.getTableWidth();
        this.table.setWidth(tableWidth);

        CTTbl ctTbl = this.table.getCTTbl();
        CTTblPr tblPr = ctTbl.getTblPr() == null ? ctTbl.addNewTblPr() : ctTbl.getTblPr();
        CTTblLayoutType tblLayout = tblPr.getTblLayout() == null ? tblPr.addNewTblLayout() : tblPr.getTblLayout();
        tblLayout.setType(STTblLayoutType.Enum.forString("fixed"));
        // 设置单元格边距
        CTTblCellMar ctTblCellMar = tblPr.addNewTblCellMar();
        float convertRate = 0.00176366843033509700176366843034F;
        float allValue = 0.1F;
//        CTTblWidth ctTblWidthTop = WordStyleUtil.createCTTblWidth(allValue / convertRate, "dxa");
//        ctTblCellMar.setTop(ctTblWidthTop);
        CTTblWidth ctTblWidthLeft = WordStyleUtil.createCTTblWidth(allValue / convertRate, "dxa");
        ctTblCellMar.setLeft(ctTblWidthLeft);
//        CTTblWidth ctTblWidthBottom = WordStyleUtil.createCTTblWidth(allValue / convertRate, "dxa");
//        ctTblCellMar.setBottom(ctTblWidthBottom);
        CTTblWidth ctTblWidthRight = WordStyleUtil.createCTTblWidth(allValue / convertRate, "dxa");
        ctTblCellMar.setRight(ctTblWidthRight);
    }

    @Override
    protected AbsWriteStyle<?> createWriteStyle() {
        return new WordWriteStyle();
    }

    @Override
    protected void setCellValue(int y, int x, Object value, AbsWriteStyle<?> writeStyle) {
        String font = writeStyle.getFont();
        boolean oneRow = writeStyle.isOneRow();
        boolean fontBold = writeStyle.isFontBold();
        String fontColor = writeStyle.getFontColor();
        XWPFTableRow row = this.table.getRow(y) == null ? this.table.createRow() : this.table.getRow(y);
        String valueStr = value.toString();
        XWPFTableCell cell = row.getCell(x) == null ? row.createCell() : row.getCell(x);
        this.setBorders(cell, writeStyle);
        if (oneRow) {
            XWPFParagraph paragraph = cell.getParagraphArray(0);
            if (paragraph == null) {
                paragraph = cell.addParagraph();
            }
            WordCell.setCellCenter(cell, paragraph, writeStyle.getVertical(), writeStyle.getHorizontal());
            XWPFRun run = paragraph.createRun();
            run.setBold(fontBold);
            run.setColor(fontColor);
            run.setText(valueStr);
            WordStyleUtil.setFont(run, font);
            return;
        }
        String[] titleSplit = valueStr.split("\n");
        for (int i = 0; i < titleSplit.length; i++) {
            String item = titleSplit[i];
            if (StrUtil.isBlank(item)) {
                continue;
            }
            XWPFParagraph paragraph = cell.getParagraphArray(i);
            if (paragraph == null) {
                paragraph = cell.addParagraph();
            }
            WordCell.setCellCenter(cell, paragraph, writeStyle.getVertical(), writeStyle.getHorizontal());
            XWPFRun run = paragraph.createRun();
            run.setBold(fontBold);
            run.setColor(fontColor);
            run.setText(item);
            WordStyleUtil.setFont(run, font);
        }
    }

    private void setBorders(XWPFTableCell cell, AbsWriteStyle<?> writeStyle) {
        CTTc ctTc = cell.getCTTc();
        CTTcPr tcPr = ctTc.getTcPr() == null ? ctTc.addNewTcPr() : ctTc.getTcPr();
        CTTcBorders ctTcBorders = tcPr.addNewTcBorders();

        CTBorder topBorder = ctTcBorders.addNewTop();
        BorderStyle borderTop = writeStyle.getBorderTop();
        topBorder.setVal(STBorder.SINGLE);
        topBorder.setSz(BigInteger.valueOf(6 * borderTop.getCode()));
        topBorder.setColor("000000");

        CTBorder rightBorder = ctTcBorders.addNewRight();
        BorderStyle borderRight = writeStyle.getBorderRight();
        rightBorder.setVal(STBorder.SINGLE);
        rightBorder.setSz(BigInteger.valueOf(6 * borderRight.getCode()));
        rightBorder.setColor("000000");

        CTBorder bottomBorder = ctTcBorders.addNewBottom();
        BorderStyle borderBottom = writeStyle.getBorderBottom();
        bottomBorder.setVal(STBorder.SINGLE);
        bottomBorder.setSz(BigInteger.valueOf(6 * borderBottom.getCode()));
        bottomBorder.setColor("000000");

        CTBorder leftBorder = ctTcBorders.addNewLeft();
        BorderStyle borderLeft = writeStyle.getBorderLeft();
        leftBorder.setVal(STBorder.SINGLE);
        leftBorder.setSz(BigInteger.valueOf(6 * borderLeft.getCode()));
        leftBorder.setColor("000000");
    }

    @Override
    public int writeData(int rowIndex, RowHandleResult<OUT_TYPE> rowHandleResult) {
        Collection<List<DrawFieldInfo>> drawOrderFieldInfoMapValues = this.drawXDrawFieldInfoMap.values();
        XWPFTableRow row = table.createRow();
        WordCell wordCell = new WordCell();
        OUT_TYPE data = rowHandleResult.getData();
        for (List<DrawFieldInfo> drawFieldInfoList : drawOrderFieldInfoMapValues) {
            int drawX = drawFieldInfoList.iterator().next().getDrawX();
            XWPFTableCell cell = row.getCell(drawX) == null ? row.createCell() : row.getCell(drawX);
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
                        WordCell childWordCell = new WordCell();
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
                AbsWriteStyle<?> dataColStyle = drawFieldInfo.createDataColStyle(this::createWriteStyle);
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
        int startX = mergePosition.getStartX();
        mergeCellsVertically(this.table, startX, startY, endY);
    }

    @Override
    public void write(OutputStream outputStream) throws IOException {
        this.document.write(outputStream);
    }

    @Override
    public int getTableWidth() {
        return this.mainWriteConfig.getDocumentWidthNotMargin();
    }

    @Override
    public void close() {
        try {
            super.close();
            document.close();
        } catch (IOException ioException) {
            ioException.printStackTrace();
        }
    }

    @Override
    protected void setColumnWidth(int y, int x, float width, boolean gridCol) {
        if (gridCol) {
            CTTbl ctTbl = this.table.getCTTbl();
            CTTblGrid tblGrid = ctTbl.getTblGrid() == null ? ctTbl.addNewTblGrid() : ctTbl.getTblGrid();
            CTTblGridCol ctTblGridCol = tblGrid.addNewGridCol();
            ctTblGridCol.setW(BigDecimal.valueOf(((long) width)));
            return;
        }
        XWPFTableRow row = this.table.getRow(y) == null ? this.table.createRow() : this.table.getRow(y);
        XWPFTableCell cell = row.getCell(x) == null ? row.createCell() : row.getCell(x);
        cell.setWidthType(TableWidthType.DXA);
        cell.setWidth(width <= -1F ? "auto" : (((int) width) + ""));
    }

    @Override
    protected void setRowHeight(int y, float height) {
        XWPFTableRow row = this.table.getRow(y) == null ? this.table.createRow() : this.table.getRow(y);
        row.setHeight((int) height);
    }

    @Override
    protected void autoRowHeight(int y, boolean autoHeight, float defaultHeight, float heightRate) {
        XWPFTableRow row = this.table.getRow(y);
        List<XWPFTableCell> tableCells = row.getTableCells();
        long maxCharSum = 0;
        int tableCellsSize = tableCells.size();
        for (int colIx = 0; colIx < tableCellsSize; colIx++) {
            XWPFTableCell tableCell = tableCells.get(colIx);
            String cellText = tableCell.getText();
            List<DrawFieldInfo> drawFieldInfoList = this.drawXDrawFieldInfoMap.get(colIx);
            float contentHeightSum = 0F;
            for (DrawFieldInfo drawFieldInfo : drawFieldInfoList) {
                contentHeightSum += calcHeight(colIx, tableCell, cellText, drawFieldInfo, heightRate);
                if (super.dataTitleRowPosition == y) {
                    break;
                }
            }
            maxCharSum = Math.max(maxCharSum, (long) (contentHeightSum));
        }
        row.setHeight((int) maxCharSum);
    }

    private float calcHeight(int colIx, XWPFTableCell tableCell, String cellText, DrawFieldInfo drawFieldInfo, float heightRate) {
        if (drawFieldInfo.isCustom()) {
            return 0;
        }
        AbsWriteStyle<?> writeStyle = drawFieldInfo.createDataColStyle(this::createWriteStyle);
        int lineIndentNum = writeStyle.getLineIndentNum();
        List<XWPFParagraph> xwpfParagraphList = tableCell.getParagraphs();
        float newlineSum = xwpfParagraphList.size();
        long dataCharLength = CharUtil.dataCharLength(cellText);
        if (lineIndentNum == -1 || lineIndentNum > 0) {
            dataCharLength += ((lineIndentNum == -1) ? 4 : lineIndentNum);
        }
        float cellRowMaxCharCount = this.getCellRowMaxCharCount(colIx);
//            cellRowMaxCharCount -= ((dataCharLength / cellRowMaxCharCount) * 0.1F);
        float contentRow = 1F;
        if (dataCharLength >= cellRowMaxCharCount) {
            BigDecimal bigDecimal = BigDecimal.valueOf(dataCharLength / cellRowMaxCharCount);
            contentRow = bigDecimal.setScale(0, RoundingMode.UP).floatValue();
        }
        contentRow += (newlineSum / 2F);
        float fontSizePt = writeStyle.getFontSizePoints() * 2;
        //12.15F
        float rate = 0;
        if (contentRow > 20) {
            rate = 0.5F;
        } else if (contentRow > 10) {
            rate = 0.25F;
        }
        float floatValue = BigDecimal.valueOf(contentRow * rate).setScale(0, RoundingMode.DOWN).floatValue();
        float contentHeight = (((contentRow - floatValue) * heightRate) * (fontSizePt * 9.5F));
        return contentHeight;
    }

    @Override
    protected float _getCellRowMaxCharCount(int x) {
        List<DrawFieldInfo> drawFieldInfoList = this.drawXDrawFieldInfoMap.get(x);
        final DrawFieldInfo drawFieldInfo = drawFieldInfoList.iterator().next();
        AbsWriteStyle<?> dataColStyle = drawFieldInfo.createDataColStyle(this::createWriteStyle);
        float fontSize = dataColStyle.getFontSizePoints() * 2;
        float width = drawFieldInfo.getWidth();
        float cellRowMaxCharCount = SizeConvertUtil.getCellRowMaxCharCount(fontSize, width);
        return cellRowMaxCharCount;
    }

    public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            CTVMerge ctvMerge = cell.getCTTc().addNewTcPr().addNewVMerge();
            if (rowIndex == fromRow) {
                ctvMerge.setVal(STMerge.RESTART);
                continue;
            }
            ctvMerge.setVal(STMerge.CONTINUE);
        }
    }
}
