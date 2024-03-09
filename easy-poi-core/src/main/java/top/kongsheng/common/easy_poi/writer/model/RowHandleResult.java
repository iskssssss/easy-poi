package top.kongsheng.common.easy_poi.writer.model;

import lombok.Data;
import lombok.ToString;

/**
 * 行转换结果
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/11/3 12:10
 */
@Data
@ToString
public class RowHandleResult<OUT_TYPE> {

    /**
     * 数据
     */
    private OUT_TYPE data;

    /**
     * 字体颜色
     */
    private String fontColor;

    public RowHandleResult(OUT_TYPE data) {
        this.data = data;
    }
}
