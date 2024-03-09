package top.kongsheng.common.easy_poi.iterator;

import top.kongsheng.common.easy_poi.reader.exception.DataImportException;
import top.kongsheng.common.easy_poi.reader.handler.ImportInfo;
import top.kongsheng.common.easy_poi.model.ReadCellInfo;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.Closeable;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * 数据链式读取接口
 *
 * @param <E> 类型
 */
public abstract class Iterator<E> implements Closeable {

    protected int lastTableIndex = 0, curTableIndex = 0;
    protected int lastRowIndex = 0, curRowIndex = 0;

    protected int tableRowTotal = -1, tableTotal = 0;

    protected final ImportInfo<?, CellRangeAddress> importInfo;
    protected final List<String> modelTitleInfoList;
    protected final Map<Integer, String> titleIndexMap;
    protected final List<ReadCellInfo> dataList = new LinkedList<>();

    protected List<CellRangeAddress> mergedRegions;

    protected Iterator(ImportInfo<?, CellRangeAddress> importInfo) {
        this.importInfo = importInfo;
        this.modelTitleInfoList = importInfo.getModelTitleInfoList();
        this.titleIndexMap = importInfo.getTitleIndexMap();
    }

    /**
     * 是否有下一行
     *
     * @return 是否有下一行
     */
    public abstract boolean hasNext();

    /**
     * 获取下一行数据
     *
     * @return 数据
     */
    public abstract E next() throws DataImportException;

    protected void setMergedRegions(List<CellRangeAddress> mergedRegions) {
        this.mergedRegions = mergedRegions;
        this.importInfo.setMergedRegions(mergedRegions);
    }

    @Override
    public void close() {
        this.mergedRegions.clear();
        this.modelTitleInfoList.clear();
        this.titleIndexMap.clear();
        this.dataList.clear();
    }
}
