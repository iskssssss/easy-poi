package top.kongsheng.common.easy_poi.reader;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.easy_poi.iterator.Iterator;
import top.kongsheng.common.easy_poi.reader.exception.DataImportException;
import top.kongsheng.common.easy_poi.reader.model.ReadResult;
import top.kongsheng.common.easy_poi.model.ReadCellInfo;
import top.kongsheng.common.easy_poi.reader.abs.AbsDataReader;
import top.kongsheng.common.easy_poi.reader.abs.AbsRowHandler;
import top.kongsheng.common.easy_poi.reader.handler.ImportInfo;
import top.kongsheng.common.easy_poi.reader.model.ErrorInfo;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 文档表格读取器
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/4 17:26
 */
public class WordTableDataReader<T> extends AbsDataReader<T, CellRangeAddress> {
    private static final String RESTART = "restart";

    private final File importExcelTemp;
    private final BufferedInputStream inputStream;
    private final XWPFDocument xwpf;
    private final WordIterator wordIterator;

    public WordTableDataReader(File importExcelTemp, AbsRowHandler<T> rowHandler) throws IOException {
        super(rowHandler);
        this.importExcelTemp = importExcelTemp;
        super.setImportFile(importExcelTemp);
        this.inputStream = FileUtil.getInputStream(importExcelTemp);
        this.xwpf = new XWPFDocument(inputStream);

        this.wordIterator = new WordIterator(this, this.xwpf);
    }

    @Override
    public Iterator<ReadData> readIterator() {
        return this.wordIterator;
    }

    @Override
    public void writeErrorInfo() {
        try (BufferedOutputStream outputStream = FileUtil.getOutputStream(importExcelTemp)) {
            List<ErrorInfo> errorInfoList = this.readResult.getErrorInfoList();
            for (ErrorInfo errorInfo : errorInfoList) {
                String tableIndexStr = errorInfo.getTableIndex();
                if (!NumberUtil.isNumber(tableIndexStr)) {
                    continue;
                }
                int tableIndex = Integer.parseInt(tableIndexStr);
                XWPFTable table = this.xwpf.getTableArray(tableIndex);
                XWPFTableRow row = table.getRow(errorInfo.getY());
                XWPFTableCell cell = row.getCell(errorInfo.getX());
                cell.removeParagraph(0);
                XWPFParagraph paragraph = cell.addParagraph();
                XWPFRun run = paragraph.createRun();
                run.setBold(true);
                run.setColor("FF0000");
                run.setText(errorInfo.getVal());
            }
            this.xwpf.write(outputStream);
            outputStream.flush();
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("[数据导入] 异常文件写入失败：" + e.getMessage());
        }
    }

    @Override
    public void close() throws IOException {
        super.close();
        if (this.wordIterator != null) {
            this.wordIterator.close();
        }
        inputStream.close();
        xwpf.close();
    }

    private static class WordIterator extends Iterator<ReadData> {
        private boolean isNextTable;
        private final List<XWPFTable> tableList;
        private List<XWPFTableRow> rows;

        private WordIterator(ImportInfo<?, CellRangeAddress> importInfo, XWPFDocument xwpf) {
            super(importInfo);
            this.tableList = xwpf.getTables();
            super.tableTotal = this.tableList.size();
            this.curTableIndex = 0;
            super.setMergedRegions(new LinkedList<>());
            this.isNextTable = true;
        }

        @Override
        public boolean hasNext() {
            return (super.curRowIndex < super.tableRowTotal) || (super.curTableIndex < super.tableTotal);
        }

        @Override
        public ReadData next() throws DataImportException {
            if (this.isNextTable) {
                if (super.curTableIndex >= tableList.size()) {
                    throw new DataImportException("数据已全部获取。");
                }
                super.lastTableIndex = super.curTableIndex;
                XWPFTable table = tableList.get(super.curTableIndex++);
                this.rows = table.getRows();
                super.tableRowTotal = this.rows.size();
                super.curRowIndex = 0;
                titleIndexMap.clear();
                if (!readTitle()) {
                    return this.next();
                }
                this.isNextTable = false;
            }
            super.dataList.forEach(ReadCellInfo::close);
            super.dataList.clear();
            this.readData();
            super.lastRowIndex = super.curRowIndex;
            super.curRowIndex++;
            if (super.curRowIndex >= super.tableRowTotal) {
                this.isNextTable = true;
            }
            return new ReadData(String.valueOf(super.lastTableIndex), super.lastRowIndex, super.dataList);
        }

        /**
         * 读取数据
         */
        private void readData() {
            XWPFTableRow xwpfTableRow = rows.get(super.curRowIndex);
            List<XWPFTableCell> tableCells = xwpfTableRow.getTableCells();
            int tableCellsSize = tableCells.size();
            for (int i = 0; i < tableCellsSize; i++) {
                XWPFTableCell tableCell = tableCells.get(i);
                Optional<CellRangeAddress> rangeAddressOptional = findCellRangeAddress(super.curRowIndex, this.rows, tableCell, i);
                rangeAddressOptional.ifPresent(this.mergedRegions::add);
                List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
                if (paragraphs.size() < 2) {
                    ReadCellInfo readCellInfo = new ReadCellInfo(tableCell.getText());
                    readCellInfo.setBackgroundColor(tableCell.getColor());
                    this.dataList.add(readCellInfo);
                    continue;
                }
                List<Object> valueList = new ArrayList<>();
                for (XWPFParagraph paragraph : paragraphs) {
                    List<XWPFRun> runs = paragraph.getRuns();
                    valueList.add(runs.stream().map(StrUtil::toString).collect(Collectors.joining("")));
                }
                ReadCellInfo readCellInfo = new ReadCellInfo(valueList);
                readCellInfo.setBackgroundColor(tableCell.getColor());
                this.dataList.add(readCellInfo);
            }
        }

        private static Optional<CellRangeAddress> findCellRangeAddress(int curRowIndex, List<XWPFTableRow> rows, XWPFTableCell tableCell, int firstCol) {
            // 跨行 restart 为分界线
            CTVMerge vMerge = tableCell.getCTTc().getTcPr().getVMerge();
            if (vMerge == null) {
                return Optional.empty();
            }
            STMerge.Enum val = vMerge.getVal();
            if (val == null || !RESTART.equals(val.toString())) {
                return Optional.empty();
            }
            CTDecimalNumber colspan = tableCell.getCTTc().getTcPr().getGridSpan();//跨列
            int colspanNum = colspan == null || colspan.getVal() == null ? 1 : colspan.getVal().intValue();
            int lastRow = 0;
            int tableRowTotal = rows.size();
            for (int i = curRowIndex + 1; i < tableRowTotal; i++) {
                XWPFTableRow xwpfTableRow = rows.get(i);
                List<XWPFTableCell> tableCells = xwpfTableRow.getTableCells();
                if (CollUtil.isEmpty(tableCells)) {
                    break;
                }
                if (tableCells.size() > firstCol) {
                    XWPFTableCell xwpfTableCell = tableCells.get(firstCol);
                    // 跨行 restart 为分界线
                    CTVMerge xwpfvMerge = xwpfTableCell.getCTTc().getTcPr().getVMerge();
                    if (xwpfvMerge == null) {
                        break;
                    }
                    STMerge.Enum xwpfVal = xwpfvMerge.getVal();
                    if (xwpfVal != null && RESTART.equals(xwpfVal.toString())) {
                        break;
                    }
                    lastRow = i;
                }
            }
            if (curRowIndex != lastRow) {
                return Optional.of(new CellRangeAddress(curRowIndex, lastRow, firstCol, firstCol + colspanNum - 1));
            }
            return Optional.empty();
        }

        /**
         * 读取标题信息
         *
         * @return 是否进入下一行
         */
        private boolean readTitle() {
            List<XWPFTableCell> tableCells = rows.get(super.curRowIndex).getTableCells();
            for (int i = 0; i < tableCells.size(); i++) {
                XWPFTableCell cell = tableCells.get(i);
                String sourceData = cell.getText().replaceAll(" ", "");
                String title = sourceData.replaceAll("\r", "").replaceAll("\n", "");
                if (!modelTitleInfoList.contains(sourceData) && !modelTitleInfoList.contains(title)) {
                    return false;
                }
                titleIndexMap.put(i, title);
            }
            super.curRowIndex++;
            ReadResult<?> readResult = this.importInfo.getReadResult();
            readResult.setFileCheck(true);
            return true;
        }
    }
}
