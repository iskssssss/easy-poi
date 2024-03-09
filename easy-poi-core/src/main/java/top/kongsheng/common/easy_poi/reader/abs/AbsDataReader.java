package top.kongsheng.common.easy_poi.reader.abs;

import top.kongsheng.common.easy_poi.iterator.Iterator;
import top.kongsheng.common.easy_poi.reader.ImportFileManager;
import top.kongsheng.common.easy_poi.reader.core.ReadResultCallback;
import top.kongsheng.common.easy_poi.reader.exception.DataImportException;
import top.kongsheng.common.easy_poi.reader.handler.ImportInfo;
import top.kongsheng.common.easy_poi.model.ReadCellInfo;

import java.io.File;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * 抽象数据读取器
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/4 17:24
 */
public abstract class AbsDataReader<T, MERGED_TYPE> extends ImportInfo<T, MERGED_TYPE> implements Runnable {
    private final AtomicBoolean readStatus = new AtomicBoolean(false);
    private final CountDownLatch countDownLatch = new CountDownLatch(1);
    private transient DataImportException dataImportException;
    private ReadResultCallback<T> readResultCallback = result -> {};

    protected final AbsRowHandler<T> rowHandler;

    protected AbsDataReader(AbsRowHandler<T> rowHandler) {
        this.rowHandler = rowHandler;
        rowHandler.init(this);
    }

    /**
     * 获取读取流
     *
     * @return 读取流
     */
    public abstract Iterator<ReadData> readIterator();

    /**
     * 写入异常信息
     */
    public abstract void writeErrorInfo();

    /**
     * 读取数据 异步
     */
    public AbsDataReader<T, MERGED_TYPE> readAsync() throws InterruptedException, DataImportException {
        return this.read(true);
    }

    /**
     * 读取数据 同步
     */
    public AbsDataReader<T, MERGED_TYPE> readSync() throws InterruptedException, DataImportException {
        return this.read(false);
    }

    public void callback(ReadResultCallback<T> readResultCallback) {
        if (readResultCallback == null) {
            return;
        }
        this.readResultCallback = readResultCallback;
    }

    /**
     * 读取数据 同步/异步
     */
    public AbsDataReader<T, MERGED_TYPE> read(boolean async) throws InterruptedException, DataImportException {
        if (readStatus.get()) {
            return this;
        }
        readStatus.set(true);
        new Thread(this).start();
        if (async) {
            return this;
        }
        countDownLatch.await();
        if (this.dataImportException != null) {
            throw this.dataImportException;
        }
        return this;
    }

    @Override
    public void run() {
        try {
            Iterator<ReadData> readIterator = this.readIterator();
            while (readIterator.hasNext()) {
                ReadData readData = readIterator.next();
                T t = rowHandler._to(readData.getTableIndex(), readData.getY(), readData.getDataList());
                if (t == null) {
                    continue;
                }
                this.readResult.addResult(t);
            }
            if (!this.readResult.getFileCheck()) {
                this.dataImportException = new DataImportException("数据导入失败，模板错误。");
            }
            // 全部读取完后的处理
            rowHandler.afterHandler(this.readResult.getResultList());
            readResultCallback.handle(this);
        } catch (Exception e) {
            this.dataImportException = new DataImportException("数据导入失败。", e);
        } finally {
            countDownLatch.countDown();
        }
    }

    /**
     * 设置导入文件
     *
     * @param file
     */
    public void setImportFile(File file) {
        this.readResult.setFileId(ImportFileManager.put(file));
    }

    public AbsRowHandler<T> getRowHandler() {
        return rowHandler;
    }

    /**
     * 根据文件id获取错误文件
     *
     * @param fileId 文件id
     * @return 错误文件
     */
    public static File findErrorImportFile(String fileId) {
        File file = ImportFileManager.find(fileId);
        return file;
    }

    protected static class ReadData {
        private final String tableIndex;
        private final int y;
        private final List<ReadCellInfo> dataList;

        public ReadData(String tableIndex, int y, List<ReadCellInfo> dataList) {
            this.tableIndex = tableIndex;
            this.y = y;
            this.dataList = dataList;
        }

        public String getTableIndex() {
            return tableIndex;
        }

        public int getY() {
            return y;
        }

        public List<ReadCellInfo> getDataList() {
            return dataList;
        }
    }
}
