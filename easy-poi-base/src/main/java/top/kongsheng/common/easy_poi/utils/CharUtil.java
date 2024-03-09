package top.kongsheng.common.easy_poi.utils;

import java.util.function.Consumer;

/**
 * 字符工具类
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/9/4 16:45
 */
public class CharUtil {

    public static long dataCharLength(String valueStr) {
        return dataCharLength(valueStr, character -> { });
    }

    public static long dataCharLength(String valueStr, Consumer<Character> consumer) {
        int valueStrLength = valueStr.length();
        long dataCharLength = 0;
        for (int i = 0; i < valueStrLength; i++) {
            char c = valueStr.charAt(i);
            consumer.accept(c);
            dataCharLength = dataCharLength + ((String.valueOf(c).matches("[^\\x00-\\xff]") || Character.charCount(c) > 2 || isChineseCharacter(c)) ? 2 : 1);
        }
        return dataCharLength;
    }

    public static boolean isChineseCharacter(char ch) {
        Character.UnicodeBlock block = Character.UnicodeBlock.of(ch);
        return block == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS
                || block == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_A
                || block == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_B
                || block == Character.UnicodeBlock.CJK_COMPATIBILITY_IDEOGRAPHS
                || block == Character.UnicodeBlock.CJK_COMPATIBILITY_IDEOGRAPHS_SUPPLEMENT;
    }
}
