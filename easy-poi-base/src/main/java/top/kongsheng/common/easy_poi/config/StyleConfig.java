package top.kongsheng.common.easy_poi.config;

import top.kongsheng.common.easy_poi.anno.PoiModel;
import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.utils.FieldUtil;
import lombok.Data;
import lombok.ToString;

/**
 * 样式配置
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/10/13 15:19
 */
@Data
@ToString
public class StyleConfig {

    /**
     * 字体
     */
    private String font;
    /**
     * 字体大小
     */
    private Float fontSizePoints;
    /**
     * 文本位置
     */
    private Integer textPosition;
    /**
     * 是否粗体
     */
    private Boolean fontBold;
    /**
     * 字体颜色
     */
    private String fontColor;
    /**
     * 水平位置
     */
    private String horizontal;
    /**
     * 垂直位置
     */
    private String vertical;

    /**
     * 边框信息
     */
    private Integer border;

    /**
     * 行缩进
     */
    private Integer lineIndentNum;

    public static StyleConfig createByPoiModel(PoiModel poiModel) {
        StyleConfig styleConfig = new StyleConfig();
        styleConfig.setFont(poiModel.font());
        styleConfig.setFontSizePoints(poiModel.fontSizePoints());
        styleConfig.setTextPosition(poiModel.textPosition());
        styleConfig.setFontBold(poiModel.fontBold());
        styleConfig.setFontColor(poiModel.fontColor());
        styleConfig.setHorizontal(poiModel.horizontal().name());
        styleConfig.setVertical(poiModel.vertical().name());
        styleConfig.setBorder(poiModel.border());
        return styleConfig;
    }

    public static StyleConfig createByPoiModelFieldWriteConfig(PoiModelField.WriteConfig writeConfig) {
        StyleConfig styleConfig = new StyleConfig();
        styleConfig.setFont(writeConfig.font());
        styleConfig.setFontSizePoints(writeConfig.fontSizePoints());
        styleConfig.setTextPosition(writeConfig.textPosition());
        styleConfig.setFontBold(writeConfig.fontBold());
        styleConfig.setFontColor(writeConfig.fontColor());
        styleConfig.setHorizontal(writeConfig.horizontal().name());
        styleConfig.setVertical(writeConfig.vertical().name());
        styleConfig.setBorder(writeConfig.border());
        styleConfig.setLineIndentNum(writeConfig.lineIndentNum());
        return styleConfig;
    }

    public boolean isFontBold() {
        return fontBold != null && fontBold;
    }

    /**
     * 更新配置
     *
     * @param styleConfig 新配置
     */
    public void update(StyleConfig styleConfig) {
        if (styleConfig == null) {
            return;
        }
        FieldUtil.updateNotNull(styleConfig.getFont(), this::setFont);
        FieldUtil.updateNotNull(styleConfig.getFontSizePoints(), this::setFontSizePoints);
        FieldUtil.updateNotNull(styleConfig.getTextPosition(), this::setTextPosition);
        FieldUtil.updateNotNull(styleConfig.getFontBold(), this::setFontBold);
        FieldUtil.updateNotNull(styleConfig.getFontColor(), this::setFontColor);
        FieldUtil.updateNotNull(styleConfig.getHorizontal(), this::setHorizontal);
        FieldUtil.updateNotNull(styleConfig.getVertical(), this::setVertical);
        FieldUtil.updateNotNull(styleConfig.getLineIndentNum(), this::setLineIndentNum);

    }

    public StyleConfig copy() {
        StyleConfig styleConfig = new StyleConfig();
        styleConfig.setFont(this.getFont());
        styleConfig.setFontSizePoints(this.getFontSizePoints());
        styleConfig.setTextPosition(this.getTextPosition());
        styleConfig.setFontBold(this.getFontBold());
        styleConfig.setFontColor(this.getFontColor());
        styleConfig.setHorizontal(this.getHorizontal());
        styleConfig.setVertical(this.getVertical());
        styleConfig.setBorder(this.getBorder());
        styleConfig.setLineIndentNum(this.getLineIndentNum());
        return styleConfig;
    }
}
