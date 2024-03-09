package top.kongsheng.common.easy_poi.model;

import top.kongsheng.common.easy_poi.config.StyleConfig;
import top.kongsheng.common.easy_poi.enums.HorizontalAlignment;
import top.kongsheng.common.easy_poi.enums.VerticalAlignment;
import top.kongsheng.common.easy_poi.utils.FieldUtil;
import lombok.Data;
import lombok.ToString;

import java.io.Serializable;

/**
 * 绘制配置信息
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/10/12 16:34
 */
@Data
@ToString
public class PoiModelFieldWriteConfig implements Serializable {
    static final long serialVersionUID = 42L;

    /**
     * 绘制字段名称
     */
    private String drawFieldKey;
    /**
     * 标题
     */
    private String title;
    /**
     * 宽度比
     */
    private Float widthRate;
    /**
     * 是否自动宽度
     */
    private Boolean autoWidth;
    /**
     * 是否自动宽度比
     *
     * @return 是否自动宽度比
     */
    private Float autoWidthRate;
    /**
     * 是否创建新样式
     */
    private Boolean newStyle;
    /**
     * 标题单元样式
     */
    private StyleConfig titleStyleConfig;
    /**
     * 数据单元样式
     */
    private StyleConfig dataStyleConfig;

    public void setTitleStyleConfig(StyleConfig titleStyleConfig, boolean mod) {
        this.titleStyleConfig = titleStyleConfig;
        if (!mod) {
            return;
        }
        this.titleStyleConfig.setFontColor("000000");
        this.titleStyleConfig.setHorizontal(HorizontalAlignment.CENTER.name());
        this.titleStyleConfig.setVertical(VerticalAlignment.CENTER.name());
        this.titleStyleConfig.setFontBold(true);
        Integer border = titleStyleConfig.getBorder();
        this.titleStyleConfig.setBorder(border + 1);
    }

    public void setDataStyleConfig(StyleConfig dataStyleConfig) {
        this.dataStyleConfig = dataStyleConfig;
    }

    public boolean isNewStyle() {
        return this.newStyle != null && this.newStyle;
    }

    public boolean isAutoWidth() {
        return this.autoWidth != null && this.autoWidth;
    }

    public void update(PoiModelFieldWriteConfig drawFieldInfo) {
        if (drawFieldInfo == null) {
            return;
        }
        FieldUtil.updateNotNull(drawFieldInfo.getTitle(), this::setTitle);
        FieldUtil.updateNotNull(drawFieldInfo.getWidthRate(), this::setWidthRate);
        FieldUtil.updateNotNull(drawFieldInfo.getAutoWidth(), this::setAutoWidth);
        FieldUtil.updateNotNull(drawFieldInfo.getAutoWidthRate(), this::setAutoWidthRate);
        FieldUtil.updateNotNull(drawFieldInfo.getNewStyle(), this::setNewStyle);
        if (this.titleStyleConfig == null) {
            this.titleStyleConfig = drawFieldInfo.getTitleStyleConfig();
        } else {
            this.titleStyleConfig.update(drawFieldInfo.getTitleStyleConfig());
        }
        if (this.dataStyleConfig == null) {
            this.dataStyleConfig = drawFieldInfo.getDataStyleConfig();
        } else {
            this.dataStyleConfig.update(drawFieldInfo.getDataStyleConfig());
        }
    }
}
