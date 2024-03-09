package top.kongsheng.common.easy_poi.enums;

/**
 * BorderStyle
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/3/22 15:33
 */
public enum BorderStyle {

    /**
     * No border (default)
     */
    NONE(0x0),

    /**
     * Thin border
     */
    THIN(0x1),

    /**
     * Medium border
     */
    MEDIUM(0x2),

    /**
     * dash border
     */
    DASHED(0x3),

    /**
     * dot border
     */
    DOTTED(0x4),

    /**
     * Thick border
     */
    THICK(0x5),

    /**
     * double-line border
     */
    DOUBLE(0x6),

    /**
     * hair-line border
     */
    HAIR(0x7),

    /**
     * Medium dashed border
     */
    MEDIUM_DASHED(0x8),

    /**
     * dash-dot border
     */
    DASH_DOT(0x9),

    /**
     * medium dash-dot border
     */
    MEDIUM_DASH_DOT(0xA),

    /**
     * dash-dot-dot border
     */
    DASH_DOT_DOT(0xB),

    /**
     * medium dash-dot-dot border
     */
    MEDIUM_DASH_DOT_DOT(0xC),

    /**
     * slanted dash-dot border
     */
    SLANTED_DASH_DOT(0xD);

    private final short code;

    private BorderStyle(int code) {
        this.code = (short)code;
    }

    public short getCode() {
        return code;
    }

    private static final BorderStyle[] _table = new BorderStyle[0xD + 1];
    static {
        for (BorderStyle c : values()) {
            _table[c.getCode()] = c;
        }
    }

    public static BorderStyle valueOf(short code) {
        return _table[code];
    }
}
