package top.kongsheng.common.easy_poi.utils;

import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.math.RoundingMode;

/**
 * WordStyleUtil
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/23 9:47
 */
public class WordStyleUtil {

    public static void setFont(XWPFRun run, String font) {
        CTFonts ctFonts = run.getCTR().addNewRPr().addNewRFonts();
        ctFonts.setEastAsia(font);
        ctFonts.setAscii(font);
        ctFonts.setCs(font);
        ctFonts.setHAnsi(font);
    }

    public static void setTableBorder(XWPFTable table, int size, String rgbColor) {
        table.setTopBorder(XWPFTable.XWPFBorderType.THICK, size, 0, rgbColor);
        table.setBottomBorder(XWPFTable.XWPFBorderType.THICK, size, 0, rgbColor);
        table.setLeftBorder(XWPFTable.XWPFBorderType.THICK, size, 0, rgbColor);
        table.setRightBorder(XWPFTable.XWPFBorderType.THICK, size, 0, rgbColor);
    }

    public static CTTblWidth createCTTblWidth(float value, String type) {
        CTTblWidth ctTblWidth = CTTblWidth.Factory.newInstance();
        BigDecimal v = BigDecimal.valueOf(value);
        ctTblWidth.setW(BigInteger.valueOf(v.setScale(0, RoundingMode.UP).longValue()));
        ctTblWidth.setType(STTblWidth.Enum.forString(type));
        return ctTblWidth;
    }
}
