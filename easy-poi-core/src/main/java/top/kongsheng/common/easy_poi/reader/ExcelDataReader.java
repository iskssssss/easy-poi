package top.kongsheng.common.easy_poi.reader;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.cell.CellUtil;
import top.kongsheng.common.easy_poi.iterator.Iterator;
import top.kongsheng.common.easy_poi.reader.abs.AbsDataReader;
import top.kongsheng.common.easy_poi.reader.abs.AbsRowHandler;
import top.kongsheng.common.easy_poi.reader.exception.DataImportException;
import top.kongsheng.common.easy_poi.reader.handler.ImportInfo;
import top.kongsheng.common.easy_poi.reader.model.ErrorInfo;
import top.kongsheng.common.easy_poi.utils.ColorUtil;
import top.kongsheng.common.easy_poi.model.ReadCellInfo;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * 表格文件内容读取器
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/4 17:26
 */
public class ExcelDataReader<T> extends AbsDataReader<T, CellRangeAddress> {
    private final File importExcelTemp;
    private final InputStream inputStream;
    private final ExcelReader excelReader;
    private final ExcelIterator excelIterator;

    public ExcelDataReader(File importExcelTemp, AbsRowHandler<T> rowHandler) {
        super(rowHandler);
        this.importExcelTemp = importExcelTemp;
        super.setImportFile(importExcelTemp);
        this.inputStream = FileUtil.getInputStream(importExcelTemp);
        ZipSecureFile.setMinInflateRatio(-1.0d);
        this.excelReader = ExcelUtil.getReader(this.inputStream);
        Sheet sheet = this.excelReader.getWorkbook().getSheetAt(0);
        Cell cell = sheet.getLastRowNum() > 0 ? sheet.getRow(0).getLastCellNum() > 0 ? sheet.getRow(0).getCell(0) : null : null;
        Object cellValue = cell != null ? CellUtil.getCellValue(cell) : null;
        if (cellValue != null) {
            this.setTopic(cellValue.toString());
        }
        this.excelIterator = new ExcelIterator(this, this.excelReader);
    }

    @Override
    public Iterator<ReadData> readIterator() {
        return this.excelIterator;
    }

    @Override
    public void writeErrorInfo() {
        Workbook workbook = this.excelReader.getWorkbook();
        try (BufferedOutputStream outputStream = FileUtil.getOutputStream(importExcelTemp)) {
            CellStyle cellStyle = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setColor(Font.COLOR_RED);
            font.setBold(true);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setFont(font);
            List<ErrorInfo> errorInfoList = this.readResult.getErrorInfoList();
            for (ErrorInfo errorInfo : errorInfoList) {
                String tableIndexStr = errorInfo.getTableIndex();
                if (!NumberUtil.isNumber(tableIndexStr)) {
                    continue;
                }
                int tableIndex = Integer.parseInt(tableIndexStr);
                Sheet sheet = workbook.getSheetAt(tableIndex);
                Row sheetRow = sheet.getRow(errorInfo.getY());
                sheetRow = sheetRow == null ? sheet.createRow(errorInfo.getY()) : sheetRow;
                Cell cell = sheetRow.getCell(errorInfo.getX());
                cell = cell == null ? sheetRow.createCell(errorInfo.getX()) : cell;
                CellStyle style = cell.getCellStyle();
                cellStyle.setBorderBottom(style.getBorderBottom());
                cellStyle.setBorderLeft(style.getBorderLeft());
                cellStyle.setBorderRight(style.getBorderRight());
                cellStyle.setBorderTop(style.getBorderTop());
                cell.setCellStyle(cellStyle);
                cell.setCellValue(errorInfo.getVal());
            }
            workbook.write(outputStream);
            outputStream.flush();
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("[数据导入] 异常文件写入失败：" + e.getMessage());
        }
    }

    @Override
    public void close() throws IOException {
        super.close();
        if (this.excelIterator != null) {
            this.excelIterator.close();
        }
        this.excelReader.close();
        this.inputStream.close();
    }

    private static class ExcelIterator extends Iterator<ReadData> {
        private int columnCount;

        private final ExcelReader excelReader;
        private final Workbook workbook;

        public ExcelIterator(ImportInfo<?, CellRangeAddress> importInfo, ExcelReader excelReader) {
            super(importInfo);
            this.excelReader = excelReader;
            this.workbook = this.excelReader.getWorkbook();
            super.tableTotal = workbook.getNumberOfSheets();
            this.curTableIndex = 0;
            Sheet sheet = workbook.getSheetAt(this.curTableIndex);
            // issues 修复导入最后一条数据缺失问题
            this.tableRowTotal = sheet.getLastRowNum() + 1;
            super.setMergedRegions(sheet.getMergedRegions());
        }

        @Override
        public boolean hasNext() {
            return super.curRowIndex < super.tableRowTotal;
        }

        @Override
        public ReadData next() throws DataImportException {
            super.dataList.forEach(ReadCellInfo::close);
            super.dataList.clear();
            if (super.curRowIndex >= this.tableRowTotal) {
                throw new DataImportException("数据已全部获取。");
            }
            if (super.titleIndexMap.isEmpty()) {
                this.columnCount = excelReader.getColumnCount(super.curRowIndex);
            }
            int readRow = 1;
            for (int i = 0; i < columnCount; i++) {
                CellRangeAddress merged = this.checkMerged(i, super.curRowIndex);
                if (merged != null && i == 0) {
                    int firstRow = merged.getFirstRow() + 1;
                    int lastRow = merged.getLastRow() + 1;
                    readRow = (lastRow - firstRow) + 1;
                }
                if (readRow > 1 && merged == null) {
                    List<Object> valueList = new LinkedList<>();
                    int endIndex = super.curRowIndex + readRow;
                    for (int y = super.curRowIndex; y < endIndex; y++) {
                        Object cellValue = excelReader.readCellValue(i, y);
                        valueList.add(cellValue);
                    }
                    ReadCellInfo readCellInfo = new ReadCellInfo(valueList);
                    readCellInfo.setBackgroundColor(ColorUtil.getColor(this.workbook, excelReader.getCell(i, super.curRowIndex)));
                    super.dataList.add(readCellInfo);
                    continue;
                }
                Object cellValue = excelReader.readCellValue(i, super.curRowIndex);
                ReadCellInfo readCellInfo = new ReadCellInfo(cellValue);
                readCellInfo.setBackgroundColor(ColorUtil.getColor(this.workbook, excelReader.getCell(i, super.curRowIndex)));
                super.dataList.add(readCellInfo);
            }
            super.lastRowIndex = super.curRowIndex;
            super.curRowIndex = super.curRowIndex + readRow;
            if (titleIndexMap.isEmpty()) {
                this.checkTitle(super.dataList);
                columnCount = titleIndexMap.size();
                return this.next();
            }
            return new ReadData(String.valueOf(super.lastTableIndex), super.lastRowIndex, super.dataList);
        }

        /**
         * 校验是否是标题行
         *
         * @param dataList 数据
         */
        private void checkTitle(List<ReadCellInfo> dataList) {
            for (int i = 0; i < dataList.size(); i++) {
                ReadCellInfo readCellInfo = dataList.get(i);
                Object data = readCellInfo.getValue();
                if (data instanceof List) {
                    data = ((List<?>) data).stream().map(StrUtil::toString).filter(StrUtil::isNotBlank).collect(Collectors.joining());
                }
                String title = StrUtil.toString(data).replaceAll(" ", "").replaceAll("\r", "").replaceAll("\n", "");
                if (!modelTitleInfoList.contains(title)) {
                    titleIndexMap.clear();
                    return;
                }
                titleIndexMap.put(i, title);
            }
            this.importInfo.getReadResult().setFileCheck(true);
        }

        /**
         * 校验是否是合并单元格
         *
         * @param xIndex 纵坐标
         * @param yIndex 横坐标
         * @return 合并信息
         */
        private CellRangeAddress checkMerged(int xIndex, int yIndex) {
            for (CellRangeAddress mergedRegion : mergedRegions) {
                int firstRow = mergedRegion.getFirstRow();
                int lastRow = mergedRegion.getLastRow();
                int firstColumn = mergedRegion.getFirstColumn();
                int lastColumn = mergedRegion.getLastColumn();
                if ((firstRow <= yIndex && yIndex <= lastRow) && (firstColumn <= xIndex && xIndex <= lastColumn)) {
                    return mergedRegion;
                }
            }
            return null;
        }
    }
}
