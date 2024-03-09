package top.kongsheng.common.easy_poi.anno;

import top.kongsheng.common.easy_poi.value.verify.FieldVerify;
import top.kongsheng.common.easy_poi.value.handle.ValueHandler;
import top.kongsheng.common.easy_poi.value.convert.WriteValueConvert;
import top.kongsheng.common.easy_poi.value.convert.WriteValueConvertDefault;
import top.kongsheng.common.easy_poi.enums.HorizontalAlignment;
import top.kongsheng.common.easy_poi.enums.VerticalAlignment;

import java.lang.annotation.*;

/**
 * PoiModelField
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/4 12:52
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface PoiModelField {

    /**
     * 标题列表
     *
     * @return 标题列表
     */
    String[] value();

    /**
     * 横坐标
     * <p>-1：不作为导入字段</p>
     *
     * @return 横坐标
     */
    int x();

    /**
     * 导入配置
     *
     * @return 导入配置
     */
    ReadConfig readConfig() default @ReadConfig();

    /**
     * 导出配置
     *
     * @return 导出配置
     */
    WriteConfig writeConfig() default @WriteConfig();

    @Retention(RetentionPolicy.RUNTIME)
    @Target({})
    @interface ReadConfig {

        /**
         * 是否必须
         *
         * @return 是否必须
         */
        boolean required() default false;

        /**
         * 字段转换器
         *
         * @return 字段转换器
         */
        Class<? extends ValueHandler> typeHandler() default ValueHandler.ValueHandlerDefault.class;

        /**
         * 字段校验器
         *
         * @return 字段校验器
         */
        Class<? extends FieldVerify> fieldVerify() default FieldVerify.FieldVerifyDefault.class;
    }

    @Retention(RetentionPolicy.RUNTIME)
    @Target({})
    @interface WriteConfig {
        /**
         * 导出标题下标
         *
         * @return 导出标题下标
         */
        int titleGetIndex() default 0;

        /**
         * 导出数据横坐标
         *
         * @return 横坐标
         */
        int x() default -1;

        /**
         * 同一单元格时的绘制顺序
         *
         * @return 同一单元格时的绘制顺序
         */
        int drawSort() default 1;

        /**
         * 写入类型
         * 1.覆盖
         * 3.插入
         * 2.追加
         * @return 写入类型
         */
        int writeType() default 1;

        boolean nested() default false;

        /**
         * 是否换行绘制
         *
         * @return 是否换行绘制
         */
        boolean newLine() default false;

        /**
         * 是否使用新的样式对象
         *
         * @return 是否使用新的样式对象
         */
        boolean newStyle() default false;

        /**
         * 数据处理器
         *
         * @return 数据处理器
         */
        Class<? extends ValueHandler> typeHandler() default ValueHandler.ValueHandlerDefault.class;

        /**
         * 时间默认导出格式
         *
         * @return 时间默认导出格式
         */
        String dateFormat() default "yyyy-MM-dd HH:mm:ss";

        /**
         * 是否存在孩子节点
         *
         * @return 是否存在孩子节点
         */
        boolean child() default false;

        /**
         * 列宽比
         *
         * @return 列宽比
         */
        float widthRate() default 0.1F;

        /**
         * 是否自动宽度
         *
         * @return 是否自动宽度
         */
        boolean autoWidth() default true;

        /**
         * 是否自动宽度比
         *
         * @return 是否自动宽度比
         */
        float autoWidthRate() default -1;

        /**
         * 是否列合并
         *
         * @return 是否列合并
         */
        boolean merge() default false;

        /**
         * 合并依据（字段列表）
         *
         * @return 合并依据（字段列表）
         */
        String[] mergeParentFieldNames() default {};

        /**
         * 导出写入器
         *
         * @return 导出写入器
         */
        Class<? extends WriteValueConvert> writeValueConvert() default WriteValueConvertDefault.class;

        /**
         * 字体
         */
        String font() default "宋体";

        /**
         * 字体大小
         */
        float fontSizePoints() default 12F;

        /**
         * 文本上下位置
         */
        int textPosition() default 32;

        /**
         * 是否粗体
         */
        boolean fontBold() default false;

        /**
         * 字体颜色
         */
        String fontColor() default "000000";

        /**
         * 边框大小 0：无边框
         */
        int border() default 1;

        /**
         * 文本横向位置
         *
         * @return 文本横向位置
         */
        HorizontalAlignment horizontal() default HorizontalAlignment.LEFT;

        /**
         * 文本列位置
         *
         * @return 文本列位置
         */
        VerticalAlignment vertical() default VerticalAlignment.CENTER;

        /**
         * 行缩进值
         *
         * @return 行缩进值
         */
        int lineIndentNum() default 0;
    }
}
