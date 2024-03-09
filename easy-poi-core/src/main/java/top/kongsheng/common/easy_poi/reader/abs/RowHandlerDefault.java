package top.kongsheng.common.easy_poi.reader.abs;

/**
 * RowHandlerDefault
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/26 14:59
 */
public class RowHandlerDefault<T> extends AbsRowHandler<T> {

    private final Class<T> importModelClass;

    public RowHandlerDefault(Class<T> importModelClass) {
        this.importModelClass = importModelClass;
    }

    @Override
    protected Class<T> importModelClass() {
        return this.importModelClass;
    }
}
