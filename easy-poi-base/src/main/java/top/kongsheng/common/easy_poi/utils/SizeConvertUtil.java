package top.kongsheng.common.easy_poi.utils;

import cn.hutool.core.util.StrUtil;

import java.math.BigDecimal;
import java.math.RoundingMode;

/**
 * 尺寸转换工具类
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/10/10 16:46
 */
public class SizeConvertUtil {

    public static float getCellRowMaxCharCount(float fontSize, float width) {
        // 像素 Points = dxa / 20
        // 英寸 Inches = Points / 72
        // 厘米 Centimeters = Inches * 2.54
        // 72pt = 25.4mm
        // 0.35277777777777777777777777777778
        // 1cm = 360000 EMUs
        float mm = pt2mm(fontSize);
        float dxa = mm2dxa(mm);
        float rowCharCount = width / dxa;
        //double rowCharCount = BigDecimal.valueOf((width / (fontSize * 4)) - 3).setScale(0, RoundingMode.UP).intValue();
        if (rowCharCount <= 0) {
            return 0;
        }
        return rowCharCount;
    }

    public static float getCellContentRow(String text, float fontSize, float width) {
        if (StrUtil.isEmpty(text)) {
            return 0;
        }
        float cellRowMaxCharCount = getCellRowMaxCharCount(fontSize, width);
        long dataCharLength = CharUtil.dataCharLength(text);
        if (dataCharLength >= cellRowMaxCharCount) {
            BigDecimal bigDecimal = BigDecimal.valueOf(dataCharLength / cellRowMaxCharCount);
            return bigDecimal.setScale(0, RoundingMode.UP).floatValue();
        }
        return 1F;
    }

    public static float getCellContentRow(String text, float cellRowMaxCharCount) {
        if (StrUtil.isEmpty(text)) {
            return 0;
        }
        long dataCharLength = CharUtil.dataCharLength(text);
        if (dataCharLength >= cellRowMaxCharCount && cellRowMaxCharCount != 0) {
            BigDecimal bigDecimal = BigDecimal.valueOf(dataCharLength / cellRowMaxCharCount);
            return bigDecimal.setScale(0, RoundingMode.UP).floatValue();
        }
        return 1F;
    }

    public static float pt2dxa(float pt) {
        float mm = pt2mm(pt);
        float dxa = mm2dxa(mm);
        return dxa;
    }

    public static float pt2mm(float pt) {
        return (pt / 4F) * 0.35277777777777777777777777777778F;
    }

    public static float mm2dxa(float mm) {
        return cm2dxa(mm / 10F);
    }

    public static float cm2dxa(float cm) {
        return (cm * 1440F) / 2.54F;
    }

    public static float dxa2cm(float dxa) {
        return (dxa / 1440F) * 2.54F;
    }
}
