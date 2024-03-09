package top.kongsheng.common.easy_poi.utils;

import org.apache.poi.hssf.record.PaletteRecord;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * poi颜色工具类
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/5/5 15:34
 */
public class ColorUtil {
    private static final PaletteRecord PALETTE_RECORD = new PaletteRecord();

    /**
     * 获取指定单元格的颜色
     *
     * @param workbook 工作簿
     * @param cell     单元格
     * @return 颜色
     */
    public static String getColor(Workbook workbook, Cell cell) {
        if (cell == null) {
            return null;
        }
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle instanceof XSSFCellStyle) {
            XSSFCellStyle xssfCellStyle = (XSSFCellStyle) cellStyle;
            byte[] brgb = getRGB(xssfCellStyle.getFillBackgroundColorColor(), xssfCellStyle.getFillBackgroundXSSFColor(), xssfCellStyle.getFillForegroundColorColor(), xssfCellStyle.getFillForegroundXSSFColor());
            if (brgb == null) {
                short fillBackgroundColor = xssfCellStyle.getFillBackgroundColor();
                brgb = PALETTE_RECORD.getColor(fillBackgroundColor + (short) 0x8);
                if (brgb == null) {
                    short fillForegroundColor = xssfCellStyle.getFillForegroundColor();
                    brgb = PALETTE_RECORD.getColor(fillForegroundColor + (short) 0x8);
                    if (brgb == null) {
                        return null;
                    }
                }
            }
            return String.format("%02X", brgb[0]) + String.format("%02X", brgb[1]) + String.format("%02X", brgb[2]);
        }
        if (cellStyle instanceof HSSFCellStyle) {
            HSSFPalette palette = ((HSSFWorkbook) workbook).getCustomPalette();
            short[] srgb = getRGB(palette, cellStyle.getFillForegroundColor(), cellStyle.getFillBackgroundColor());
            if (srgb == null) {
                return null;
            }
            return String.format("%02X", srgb[0]) + String.format("%02X", srgb[1]) + String.format("%02X", srgb[2]);
        }
        return null;
    }

    private static byte[] getRGB(XSSFColor... xssfColors) {
        byte[] temp = null;
        for (XSSFColor xssfColor : xssfColors) {
            if (xssfColor == null) {
                continue;
            }
            byte[] brgb = xssfColor.getRGB();
            if (brgb == null) {
                continue;
            }
            if ((brgb[0] == 0 && brgb[1] == 0 && brgb[2] == 0) || (brgb[0] == -1 && brgb[1] == -1 && brgb[2] == -1)) {
                temp = brgb;
                continue;
            }
            return brgb;
        }
        return temp;
    }

    private static short[] getRGB(HSSFPalette palette, short... colorIndexList) {
        short[] temp = null;
        for (short colorIndex : colorIndexList) {
            HSSFColor hssfcolor = palette.getColor(colorIndex);
            if (hssfcolor == null) {
                continue;
            }
            short[] srgb = hssfcolor.getTriplet();
            if (srgb == null) {
                continue;
            }
            if ((srgb[0] == 0 && srgb[1] == 0 && srgb[2] == 0) || (srgb[0] == -1 && srgb[1] == -1 && srgb[2] == -1) ||
                    (srgb[0] == 255 && srgb[1] == 255 && srgb[2] == 255)) {
                temp = srgb;
                continue;
            }
            return srgb;
        }
        return temp;
    }
}
