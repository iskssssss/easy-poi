package top.kongsheng.common.easy_poi.writer;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.easy_poi.model.WordCell;
import top.kongsheng.common.easy_poi.utils.WordStyleUtil;
import top.kongsheng.common.easy_poi.anno.PoiModel;
import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.writer.abs.AbsDataWriter;
import top.kongsheng.common.easy_poi.writer.abs.AbsWriterRowHandler;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigInteger;
import java.util.*;

/**
 * 数据导出工具类
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/23 20:26
 */
public class DataWriteUtil {

    public static <IN_TYPE, OUT_TYPE> void export(
            AbsWriterRowHandler<IN_TYPE, OUT_TYPE> rowHandler,
            List<String> customizeField,
            Boolean exclude,
            List<String> systemField,
            OutputStream outputStream,
            boolean excel
    ) {
        try (AbsDataWriter<IN_TYPE, OUT_TYPE> dataWriter =
                     createDataWriter(excel, rowHandler, customizeField, exclude, systemField, outputStream)) {
            dataWriter.run();
        } catch (IOException ioException) {
            ioException.printStackTrace();
        }
    }

    public static <IN_TYPE, OUT_TYPE> AbsDataWriter<IN_TYPE, OUT_TYPE> createDataWriter(
            boolean excel,
            AbsWriterRowHandler<IN_TYPE, OUT_TYPE> rowHandler,
            List<String> customizeField,
            Boolean exclude,
            List<String> systemField,
            OutputStream outputStream
    ) throws IOException {
        AbsDataWriter<IN_TYPE, OUT_TYPE> dataWriter = excel ? new ExcelDataWriter<>(rowHandler) : new WordDataWriter<>(rowHandler);
        dataWriter
                .setOutputStream(outputStream)
                .addCustomizeFieldList(customizeField)
                .setExclude(exclude)
                .addSystemFieldList(systemField);
        return dataWriter;
    }

    public static final String EXPORT_DIR = FileUtil.getTmpDir().getPath() + File.separator + "EXPORT_DIR" + File.separator;

    public static File mkdirExportDir() {
        String exportPath = EXPORT_DIR + IdUtil.fastSimpleUUID();
        return FileUtil.mkdir(exportPath);
    }

    /**
     * 获取导入模板
     *
     * @return 导入模板
     */
    public static void getImportTemplate(
            OutputStream outputStream,
            Class<?> importClass,
            boolean excel,
            String title,
            String note,
            Map<Integer, Collection<String>> listOfValuesMap
    ) throws IOException {
        if (excel) {
            genImportTemplateForExcel(outputStream, importClass, title, note, listOfValuesMap);
            return;
        }
        genImportTemplateForWord(outputStream, importClass, title, note);
    }

    /**
     * 获取导入模板
     *
     * @return 导入模板
     */
    public static void genImportTemplateForWord(OutputStream outputStream, Class<?> importClass, String title, String note) throws IOException {
        PoiModel poiModel = importClass.getAnnotation(PoiModel.class);
        XWPFDocument document = new XWPFDocument();
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        boolean setPgSz = sectPr.isSetPgSz();
        CTPageSz pgsz = setPgSz ? sectPr.getPgSz() : sectPr.addNewPgSz();
        pgsz.setW(BigInteger.valueOf(poiModel.documentWidth()));
        pgsz.setH(BigInteger.valueOf(poiModel.documentHeight()));

        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(title);
        run.setTextPosition(48);
        run.setFontSize(24);
        run.setBold(true);
        WordStyleUtil.setFont(run, "宋体");
        paragraph.setAlignment(ParagraphAlignment.CENTER);

        if (StrUtil.isNotEmpty(note)) {
            String[] notes = note.split("\n");
            for (String item : notes) {
                XWPFParagraph noteParagraph = document.createParagraph();
                XWPFRun noteRun = noteParagraph.createRun();
                noteRun.setText(item);
                noteRun.setTextPosition(18);
                noteRun.setFontSize(10);
                WordStyleUtil.setFont(noteRun, "宋体");
                noteParagraph.setAlignment(ParagraphAlignment.LEFT);
            }
        }

        XWPFTable table = document.createTable();
        table.setTableAlignment(TableRowAlign.CENTER);
        WordStyleUtil.setTableBorder(table, 16, "000000");

        Field[] declaredFields = importClass.getDeclaredFields();
        Set<Integer> createIndexSet = new HashSet<>();
        Map<Integer, String> cellWidthMap = new HashMap<>();
        for (Field declaredField : declaredFields) {
            PoiModelField fieldAnnotation = declaredField.getAnnotation(PoiModelField.class);
            if (fieldAnnotation == null) {
                continue;
            }
            PoiModelField.WriteConfig writeConfig = fieldAnnotation.writeConfig();
            int x = fieldAnnotation.x();
            if (createIndexSet.contains(x) || x == -1) {
                continue;
            }
            createIndexSet.add(x);
            int titleGetIndex = writeConfig.titleGetIndex();
            String cellTitle = fieldAnnotation.value()[titleGetIndex];
            float width = (23800F * writeConfig.widthRate());
            XWPFTableRow row = table.getRow(0);
            row.setHeight(24 * 24);
            XWPFTableCell cell = row.getCell(x) == null ? row.createCell() : row.getCell(x);
            cell.setWidthType(TableWidthType.DXA);
            String widthValue = width == -1 ? "auto" : (((int) width) + "");
            cellWidthMap.put(x, widthValue);
            cell.setWidth(widthValue);
            String[] titleSplit = cellTitle.split("\n");
            for (int i = 0; i < titleSplit.length; i++) {
                String item = titleSplit[i];
                if (StrUtil.isBlank(item)) {
                    continue;
                }
                XWPFParagraph cellParagraph = cell.getParagraphArray(i);
                if (cellParagraph == null) {
                    cellParagraph = cell.addParagraph();
                }
                WordCell.setCellCenter(cell, cellParagraph, top.kongsheng.common.easy_poi.enums.VerticalAlignment.CENTER, top.kongsheng.common.easy_poi.enums.HorizontalAlignment.CENTER);
                XWPFRun cellRun = cellParagraph.createRun();
                cellRun.setBold(true);
                cellRun.setColor("000000");
                cellRun.setText(item);
                WordStyleUtil.setFont(cellRun, "宋体");
            }
        }
        for (int i = 1; i <= 5; i++) {
            XWPFTableRow row = table.createRow();
            row.setHeight(24 * 24);
            for (Integer x : createIndexSet) {
                XWPFTableCell cell = row.getCell(x) == null ? row.createCell() : row.getCell(x);
                String widthValue = cellWidthMap.get(x);
                cell.setWidthType(TableWidthType.DXA);
                cell.setWidth(widthValue);
            }
        }
        document.write(outputStream);
        document.close();
    }

    /**
     * 获取导入模板
     *
     * @return 导入模板
     */
    public static void genImportTemplateForExcel(
            OutputStream outputStream, Class<?> importClass, String title, String note, Map<Integer, Collection<String>> listOfValuesMap
    ) throws IOException {
        Workbook workbook = WorkbookFactory.create(true);
        Sheet sheet = workbook.createSheet(title);
        int rowIndex = 0;
        CellStyle titleCellStyle = workbook.createCellStyle();
        titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Font font = workbook.createFont();
        font.setFontHeight((short) (24 * 16));
        titleCellStyle.setFont(font);
        // 绘制大标题
        Row titleRow = createRow(sheet, rowIndex++, (short) (800));
        Cell cell = titleRow.createCell(0);
        cell.setCellValue(title);
        cell.setCellStyle(titleCellStyle);
        // 绘制说明行
        if (StrUtil.isNotEmpty(note)) {
            String[] notes = note.split("\n");
            Row noteRow = createRow(sheet, rowIndex++, (short) (400 * notes.length));
            cell = noteRow.createCell(0);
            cell.setCellValue(note);
            CellStyle noteCellStyle = workbook.createCellStyle();
            Font noteFont = workbook.createFont();
            noteFont.setFontHeight((short) (24 * 8));
            noteCellStyle.setFont(noteFont);
            cell.setCellStyle(noteCellStyle);
        }
        CellStyle cellStyle = createCellStyle(workbook, BorderStyle.MEDIUM);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Field[] declaredFields = importClass.getDeclaredFields();
        Set<Integer> createIndexSet = new HashSet<>();
        Row row = createRow(sheet, rowIndex++, (short) (650));
        for (Field declaredField : declaredFields) {
            PoiModelField fieldAnnotation = declaredField.getAnnotation(PoiModelField.class);
            if (fieldAnnotation == null) {
                continue;
            }
            PoiModelField.WriteConfig writeConfig = fieldAnnotation.writeConfig();
            int x = fieldAnnotation.x();
            if (createIndexSet.contains(x) || x == -1) {
                continue;
            }
            Collection<String> strings = listOfValuesMap.get(x);
            if (strings != null) {
                DataValidation dataValidation = getDataValidationList(sheet, rowIndex, x, 1000000, x, strings);
                dataValidation.setErrorStyle(DataValidation.ErrorStyle.WARNING);
                dataValidation.setShowErrorBox(true);
                dataValidation.createErrorBox("提示", "请选择正确的数据。");
                sheet.addValidationData(dataValidation);
            }
            createIndexSet.add(x);
            int titleGetIndex = writeConfig.titleGetIndex();
            String cellTitle = fieldAnnotation.value()[titleGetIndex];
            float width = (36000F * writeConfig.widthRate());
            cell = row.createCell(x);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(cellTitle);
            sheet.setColumnWidth(x, (int) width);
        }
        int lastCol = createIndexSet.size() - 1;
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, lastCol));
        if (StrUtil.isNotEmpty(note)) {
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, lastCol));
        }
        // 绘制数据行
        CellStyle dataCellStyle = createCellStyle(workbook, BorderStyle.THIN);
        for (int y = rowIndex; y <= rowIndex + 5; y++) {
            Row dataRow = sheet.createRow(y);
            dataRow.setHeight((short) (400));
            for (int x = 0; x <= lastCol; x++) {
                Cell dataRowCell = dataRow.createCell(x);
                dataRowCell.setCellStyle(dataCellStyle);
            }
        }
        workbook.write(outputStream);
        workbook.close();
    }

    static XSSFDataValidation getDataValidationList(Sheet sheet, int firstRow, int firstCol, int endRow, int endCol, Collection<String> strList) {
        String[] datas = strList.toArray(new String[0]);
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(((XSSFSheet) sheet));
        XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
                .createExplicitListConstraint(datas);
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
        XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint, addressList);
        return validation;
    }

    private static CellStyle createCellStyle(Workbook workbook, BorderStyle border) {
        CellStyle dataCellStyle = workbook.createCellStyle();
        dataCellStyle.setBorderTop(border);
        dataCellStyle.setBorderBottom(border);
        dataCellStyle.setBorderLeft(border);
        dataCellStyle.setBorderRight(border);
        return dataCellStyle;
    }

    private static Row createRow(Sheet sheet, int rowIndex, short height) {
        Row row = sheet.createRow(rowIndex);
        row.setHeight(height);
        return row;
    }
}
