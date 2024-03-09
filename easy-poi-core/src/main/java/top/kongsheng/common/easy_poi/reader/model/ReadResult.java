package top.kongsheng.common.easy_poi.reader.model;

import lombok.Data;
import lombok.ToString;

import java.io.Serializable;
import java.util.LinkedList;
import java.util.List;

/**
 * 导入结果信息
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/5/8 14:58
 */
@Data
@ToString
public class ReadResult<T> implements Serializable {
    private static final long serialVersionUID = 42L;

    /**
     * 编码
     */
    private int code;
    /**
     * 是否成功
     */
    private boolean success;
    /**
     * 处理结果
     */
    private final List<T> resultList = new LinkedList<>();
    /**
     * 错误数据
     */
    private final List<ErrorInfo> errorInfoList = new LinkedList<>();
    /**
     * 导入文件是否正确
     */
    private Boolean fileCheck = false;
    /**
     * 文件ID
     */
    private String fileId;

    /**
     * 添加结果
     *
     * @param t 转换结果
     */
    public void addResult(T t) {
        this.resultList.add(t);
    }

    /**
     * 添加错误信息
     *
     * @param errorInfo 错误信息
     */
    public void addErrorInfoList(ErrorInfo errorInfo) {
        if (errorInfo == null) {
            return;
        }
        this.errorInfoList.add(errorInfo);
    }

    /**
     * 是否再导入时发生错误
     *
     * @return 是否再导入时发生错误
     */
    public boolean isError() {
        return !this.errorInfoList.isEmpty();
    }
}
