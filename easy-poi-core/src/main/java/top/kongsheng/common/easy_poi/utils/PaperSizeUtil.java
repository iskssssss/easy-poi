package top.kongsheng.common.easy_poi.utils;

import org.apache.poi.ss.usermodel.PrintSetup;

/**
 * 纸张大小工具类
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/10/31 17:23
 */
public class PaperSizeUtil {

    public static PaperSizeEnum find(float documentWidth, float documentHeight, boolean landscape) {
        float height = documentHeight / 20F / 28.35F;
        float width = documentWidth / 20F / 28.35F;
        // A3 - 29.7x42 cm
        PaperSizeEnum paperSizeEnum = PaperSizeEnum.to(width, height, landscape);
        return paperSizeEnum;
    }


    public enum PaperSizeEnum {
        /**
         * A3 - 29.7x42 cm
         */
        A3(29.7F, 42F, PrintSetup.A3_PAPERSIZE),
        /**
         * A4 - 21x29.7 cm
         */
        A4(21F, 29.7F, PrintSetup.A4_PAPERSIZE);

        final float w;
        final float h;
        final short size;

        PaperSizeEnum(float w, float h, short size) {
            this.w = w;
            this.h = h;
            this.size = size;
        }

        public float getW() {
            return w;
        }

        public float getH() {
            return h;
        }

        public short getSize() {
            return size;
        }

        public boolean check(float targetW, float targetH, boolean landscape) {
            double startW = w - 0.5;
            double endW = w + 0.5;
            double startH = h - 0.5;
            double endH = h + 0.5;
            if (landscape) {
                return targetW >= startH && targetW <= endH && targetH >= startW && targetH <= endW;
            }
            return targetW >= startW && targetW <= endW && targetH >= startH && targetH <= endH;
        }

        public static PaperSizeEnum to(float documentWidth, float documentHeight, boolean landscape) {
            PaperSizeEnum[] values = PaperSizeEnum.values();
            for (PaperSizeEnum paperSizeEnum : values) {
                if (paperSizeEnum.check(documentWidth, documentHeight, landscape)) {
                    return paperSizeEnum;
                }
            }
            return null;
        }
    }
}
