package top.kongsheng.common.easy_poi.model;

import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.cell.CellUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRElt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;
import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.handler.AbsWriteStyle;
import top.kongsheng.common.easy_poi.writer.style.ExcelWriteStyle;

import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * excel 单元格
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/9/15 14:30
 */
public class ExcelCell extends top.kongsheng.common.easy_poi.model.Cell<org.apache.poi.ss.usermodel.Cell> {

    @Override
    public void initStyle(AbsWriteStyle<?> writeStyle) {
        super.initStyle(writeStyle);
        CellStyle style = (CellStyle) writeStyle.get();
        Cell cell = getCell();
        cell.setCellStyle(style);
    }

    @Override
    public void writeList(PoiModelField poiItem, List<?> cellValue) {
        String valueStr = cellValue.stream().map(StrUtil::toString).collect(Collectors.joining("\n"));
        this.writeString(poiItem, valueStr);
    }

    @Override
    public void writeString(PoiModelField poiItem, String cellValue) {
        Map<PoiModelField, DrawFieldInfo> drawFieldInfoMap = getDrawFieldInfoMap();
        DrawFieldInfo drawFieldInfo = drawFieldInfoMap.get(poiItem);
        this.beforeHandle(drawFieldInfo);
        Cell cell = getCell();
        this.setValue(drawFieldInfo, cell, cellValue);
    }

    @Override
    public void setBackgroundColor(String color) {
        Cell cell = getCell();
        CellStyle cellStyle = cell.getCellStyle();
        short index;
        switch (color) {
            case "FF0000":
                index = IndexedColors.RED.getIndex();
                break;
            case "00FF00":
                index = IndexedColors.BRIGHT_GREEN.getIndex();
                break;
            case "FFFF00":
                index = IndexedColors.YELLOW.getIndex();
                break;
            default:
                index = IndexedColors.WHITE.getIndex();
        }
        cellStyle.setFillForegroundColor(index);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(cellStyle);
        Object oldValue = CellUtil.getCellValue(cell);
        CellUtil.setCellValue(cell, oldValue);
    }

    @Override
    public void setFontColor(String color) {

    }

    private void beforeHandle(DrawFieldInfo drawFieldInfo) {
        Cell cell = getCell();
        int writeType = drawFieldInfo.getWriteType();
        if (writeType == 1) {
            cell.setCellValue("");
        }
    }

    private void setValue(DrawFieldInfo drawFieldInfo, Cell cell, String value) {
        XSSFRichTextString oldRichTextString = (XSSFRichTextString) cell.getRichStringCellValue();
        boolean oldContentEmpty = "".equals(oldRichTextString.getString());
        int lineIndentNum = getLineIndentNum(value);
        String firstStr = "";
        if (lineIndentNum > 0) {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < lineIndentNum; i++) {
                sb.append(" ");
            }
            firstStr = sb.toString();
        }
        String newValue = firstStr + value;
        if (drawFieldInfo.isNewLine() && !oldContentEmpty) {
            newValue = "\n" + newValue;
        }
        ExcelWriteStyle writeStyle = (ExcelWriteStyle) getWriteStyle();
        CellStyle cellStyle = writeStyle.get();
        Workbook workbook = writeStyle.getWorkbook();
        XSSFFont font = ((XSSFFont) workbook.getFontAt(cellStyle.getFontIndex()));
        // 不可使用通过getRichStringCellValue直接获取的对象需重新创建一个后将数据复制一份至新对象中，否则会出现数据异常。
        XSSFRichTextString newRichTextString = new XSSFRichTextString();
        if (!oldContentEmpty) {
            CTRst ctRst = oldRichTextString.getCTRst();
            CTRElt[] rArray = ctRst.getRArray();
            int rArrayLength = rArray.length;
            for (int i = 0; i < rArrayLength; i++) {
                CTRElt ctrElt = rArray[i];
                String oldText = ctrElt.getT();
                XSSFFont oldFont = oldRichTextString.getFontOfFormattingRun(i);
                newRichTextString.append(oldText, oldFont);
            }
        }
        newRichTextString.append(newValue, font);
        cell.setCellValue(newRichTextString);
    }
}
