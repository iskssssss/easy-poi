package top.kongsheng.common.easy_poi.model;

import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.easy_poi.utils.WordStyleUtil;
import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.enums.HorizontalAlignment;
import top.kongsheng.common.easy_poi.enums.VerticalAlignment;
import top.kongsheng.common.easy_poi.handler.AbsWriteStyle;
import top.kongsheng.common.easy_poi.model.Cell;
import top.kongsheng.common.easy_poi.model.DrawFieldInfo;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;

import java.math.BigInteger;
import java.util.List;
import java.util.Map;

/**
 * WordCell
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/9/15 14:09
 */
public class WordCell extends Cell<XWPFTableCell> {

    @Override
    public void writeList(PoiModelField poiItem, List<?> cellValue) {
        XWPFTableCell cell = getCell();
        Map<PoiModelField, DrawFieldInfo> drawFieldInfoMap = getDrawFieldInfoMap();
        DrawFieldInfo drawFieldInfo = drawFieldInfoMap.get(poiItem);
        this.beforeHandle(drawFieldInfo);
        AbsWriteStyle<?> writeStyle = getWriteStyle();
        boolean newLine = drawFieldInfo.isNewLine();
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        boolean oldContentEmpty = paragraphs.isEmpty();
        int rowIndex = 0;
        if (newLine && !oldContentEmpty) {
            rowIndex = paragraphs.size();
        }
        int startIndex = 0, endIndex = cellValue.size();
        for (int i = startIndex; i < endIndex; i++) {
            Object obj = cellValue.get(i);
            XWPFRun run = setValue(i + rowIndex, obj == null ? "" : StrUtil.toString(obj));
            WordStyleUtil.setFont(run, writeStyle.getFont());
        }
    }

    @Override
    public void writeString(PoiModelField poiItem, String cellValue) {
        AbsWriteStyle<?> writeStyle = getWriteStyle();
        Map<PoiModelField, DrawFieldInfo> drawFieldInfoMap = getDrawFieldInfoMap();
        DrawFieldInfo drawFieldInfo = drawFieldInfoMap.get(poiItem);
        this.beforeHandle(drawFieldInfo);
        XWPFTableCell cell = getCell();
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        boolean newLine = drawFieldInfo.isNewLine();
        boolean oldContentEmpty = paragraphs.isEmpty();
        if (cellValue.contains("\n")) {
            String[] splitList = cellValue.split("\n");
            int rowIndex = 0;
            if (newLine && !oldContentEmpty) {
                rowIndex = paragraphs.size();
            }
            int startIndex = 0, endIndex = splitList.length;
            for (int i = startIndex; i < endIndex; i++) {
                XWPFRun run = setValue(i + rowIndex, splitList[i]);
                WordStyleUtil.setFont(run, writeStyle.getFont());
            }
            return;
        }
        XWPFRun run = setValue(newLine ? -1 : 0, cellValue);
        WordStyleUtil.setFont(run, writeStyle.getFont());
    }

    @Override
    public void setBackgroundColor(String color) {
        XWPFTableCell cell = getCell();
        cell.setColor(color);
    }

    @Override
    public void setFontColor(String color) {

    }

    private void beforeHandle(DrawFieldInfo drawFieldInfo) {
        XWPFTableCell cell = getCell();
        if (drawFieldInfo.getWriteType() == 1) {
            List<XWPFParagraph> paragraphs = cell.getParagraphs();
            int size = paragraphs.size();
            for (int i = 0; i < size; i++) {
                cell.removeParagraph(0);
            }
        }
    }

    private XWPFRun setValue(int pos, String cellValue) {
        XWPFTableCell cell = getCell();
        AbsWriteStyle<?> writeStyle = getWriteStyle();
        XWPFParagraph paragraphArray = cell.getParagraphArray(pos);
        XWPFParagraph paragraph = paragraphArray == null ? cell.addParagraph() : paragraphArray;
        int lineIndentNum = getLineIndentNum(cellValue);
        if (lineIndentNum > 0) {
            CTPPr ctpPr = paragraph.getCTPPr();
            CTInd ctInd = ctpPr.getInd() == null ? ctpPr.addNewInd() : ctpPr.getInd();
            ctInd.setFirstLineChars(BigInteger.valueOf(lineIndentNum * 50L));
        }
        WordCell.setCellCenter(cell, paragraph, writeStyle.getVertical(), writeStyle.getHorizontal());
        XWPFRun run = paragraph.createRun();
        run.setColor(writeStyle.getFontColor());
        run.setText(cellValue);
        return run;
    }

    public static void setCellCenter(
            XWPFTableCell cell,
            XWPFParagraph paragraph,
            VerticalAlignment verticalAlignment,
            HorizontalAlignment horizontalAlignment
    ) {
        if (verticalAlignment == VerticalAlignment.CENTER) {
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        } else if (verticalAlignment == VerticalAlignment.BOTTOM) {
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.BOTTOM);
        } else if (verticalAlignment == VerticalAlignment.TOP) {
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.TOP);
        }
        if (horizontalAlignment == HorizontalAlignment.CENTER) {
            paragraph.setAlignment(ParagraphAlignment.CENTER);
        } else if (horizontalAlignment == HorizontalAlignment.LEFT) {
            paragraph.setAlignment(ParagraphAlignment.LEFT);
        } else if (horizontalAlignment == HorizontalAlignment.RIGHT) {
            paragraph.setAlignment(ParagraphAlignment.RIGHT);
        }
    }
}
