package top.kongsheng.common.easy_poi.writer.config;

import top.kongsheng.common.easy_poi.anno.PoiModel;
import top.kongsheng.common.easy_poi.config.StyleConfig;
import top.kongsheng.common.easy_poi.utils.FieldUtil;
import lombok.Data;
import lombok.ToString;

import java.io.Serializable;

/**
 * 总体导出配置
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/10/13 14:11
 */
@Data
@ToString
public class MainWriteConfig implements Serializable {
    static final long serialVersionUID = 42L;

    /**
     * 大标题
     */
    private String mainTitle;
    /**
     * 文档宽度
     */
    private Integer documentWidth;
    /**
     * 文档高度
     */
    private Integer documentHeight;
    /**
     * 是否横向
     */
    private Boolean landscape;
    /**
     * 左边距
     */
    private Integer marLeft;
    /**
     * 右边距
     */
    private Integer marRight;
    /**
     * 上边距
     */
    private Integer marTop;
    /**
     * 下边距
     */
    private Integer marBottom;
    /**
     * 标题样式
     */
    private StyleConfig mainTitleStyle;

    public static MainWriteConfig createByPoiModel(PoiModel poiModel) {
        MainWriteConfig mainWriteConfig = new MainWriteConfig();
        mainWriteConfig.setMainTitle(poiModel.value());
        mainWriteConfig.setDocumentWidth(poiModel.documentWidth());
        mainWriteConfig.setDocumentHeight(poiModel.documentHeight());
        mainWriteConfig.setLandscape(poiModel.landscape());
        mainWriteConfig.setMarLeft(poiModel.marLeft());
        mainWriteConfig.setMarRight(poiModel.marRight());
        mainWriteConfig.setMarTop(poiModel.marTop());
        mainWriteConfig.setMarBottom(poiModel.marBottom());
        StyleConfig mainTitleStyle = StyleConfig.createByPoiModel(poiModel);
        mainWriteConfig.setMainTitleStyle(mainTitleStyle);
        return mainWriteConfig;
    }

    public int getDocumentWidthNotMargin() {
        int t = marLeft + marRight;
        return documentWidth - t;
    }

    public boolean isLandscape() {
        return landscape != null && landscape;
    }

    public void setLandscape(Boolean landscape) {
        this.landscape = landscape;
        Integer documentWidth = this.getDocumentWidth(), documentHeight = this.getDocumentHeight();
        if (documentWidth == null || documentHeight == null) {
            return;
        }
        int max = Math.max(documentWidth, documentHeight);
        int min = Math.min(documentWidth, documentHeight);
        if (this.landscape) {
            this.setDocumentWidth(max);
            this.setDocumentHeight(min);
            return;
        }
        this.setDocumentWidth(min);
        this.setDocumentHeight(max);
    }

    /**
     * 更新配置
     *
     * @param mainWriteConfig 新配置
     */
    public void update(MainWriteConfig mainWriteConfig) {
        if (mainWriteConfig == null) {
            return;
        }
        FieldUtil.updateNotNull(mainWriteConfig.getMainTitle(), this::setMainTitle);
        FieldUtil.updateNotNull(mainWriteConfig.getDocumentWidth(), this::setDocumentWidth);
        FieldUtil.updateNotNull(mainWriteConfig.getDocumentHeight(), this::setDocumentHeight);
        FieldUtil.updateNotNull(mainWriteConfig.getLandscape(), this::setLandscape);

        FieldUtil.updateNotNull(mainWriteConfig.getMarLeft(), this::setMarLeft);
        FieldUtil.updateNotNull(mainWriteConfig.getMarRight(), this::setMarRight);
        FieldUtil.updateNotNull(mainWriteConfig.getMarTop(), this::setMarTop);
        FieldUtil.updateNotNull(mainWriteConfig.getMarBottom(), this::setMarBottom);
        FieldUtil.updateNotNull(mainWriteConfig.getLandscape(), this::setLandscape);
        if (this.mainTitleStyle == null) {
            this.mainTitleStyle = mainWriteConfig.getMainTitleStyle();
            return;
        }
        this.mainTitleStyle.update(mainWriteConfig.getMainTitleStyle());
    }
}
