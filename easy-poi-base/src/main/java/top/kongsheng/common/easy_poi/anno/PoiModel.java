package top.kongsheng.common.easy_poi.anno;

import top.kongsheng.common.easy_poi.enums.HorizontalAlignment;
import top.kongsheng.common.easy_poi.enums.VerticalAlignment;

import java.lang.annotation.*;

/**
 * 导出注解
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/5 14:20
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface PoiModel {

    /**
     * 标题名称
     */
    String value() default "";

    /**
     * 大标题位置
     *
     * @return 大标题位置
     */
    int mainTitleRowIndex() default 0;

    /**
     * 标题行位置
     *
     * @return 标题行位置
     */
    int titleRowIndex() default 1;

    /**
     * 表格行高度 -1：自适应
     * <p>实际设置宽度：rowHeight * 36</p>
     *
     * @return 表格行高度
     */
    int rowHeight() default -1;

    /**
     * 文档宽度
     *
     * @return 文档宽度
     */
    int documentWidth() default 11907;

    /**
     * 文档高度
     *
     * @return 文档高度
     */
    int documentHeight() default 16840;

    boolean landscape() default true;

    /**
     * 左边距
     */
    int marLeft() default 1000;

    /**
     * 右边距
     */
    int marRight() default 1000;

    /**
     * 上边距
     */
    int marTop() default 1000;

    /**
     * 下边距
     */
    int marBottom() default 1000;

    /**
     * 字体
     */
    String font() default "宋体";

    /**
     * 字体大小
     */
    float fontSizePoints() default 16F;

    /**
     * 文本上下位置
     */
    int textPosition() default 32;

    /**
     * 是否粗体
     */
    boolean fontBold() default true;

    /**
     * 字体颜色
     */
    String fontColor() default "000000";

    /**
     * 垂直位置
     */
    VerticalAlignment vertical() default VerticalAlignment.CENTER;

    /**
     * 水平位置
     */
    HorizontalAlignment horizontal() default HorizontalAlignment.CENTER;

    /**
     * 边框大小 0：无边框
     */
    int border() default 0;
}
