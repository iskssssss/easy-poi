package top.kongsheng.common.easy_poi.reader;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.easy_poi.reader.abs.AbsDataReader;
import top.kongsheng.common.easy_poi.reader.abs.AbsRowHandler;
import top.kongsheng.common.easy_poi.reader.abs.RowHandlerDefault;
import top.kongsheng.common.easy_poi.reader.exception.DataImportException;
import top.kongsheng.common.easy_poi.model.ReadCellInfo;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.Collection;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * 数据读取工具类
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/3 17:28
 */
public final class DataReadUtil {


    /**
     * [Word] 读取信息
     *
     * @param multipartFile 文件
     * @param rowHandler    行处理器
     * @param <T>           类型
     * @return 数据读取器
     * @throws IOException 异常
     */
    public static <T> AbsDataReader<T, ?> wordReadSync(MultipartFile multipartFile, AbsRowHandler<T> rowHandler) throws IOException, DataImportException, InterruptedException {
        return readSync(multipartFile, rowHandler, false);
    }

    /**
     * [Excel] 读取信息
     *
     * @param multipartFile 文件
     * @param rowHandler    行处理器
     * @param <T>           类型
     * @return 数据读取器
     * @throws IOException 异常
     */
    public static <T> AbsDataReader<T, ?> excelReadSync(MultipartFile multipartFile, AbsRowHandler<T> rowHandler) throws IOException, DataImportException, InterruptedException {
        return readSync(multipartFile, rowHandler, true);
    }

    /**
     * 读取表格信息信息
     *
     * @param multipartFile 文件
     * @param rowHandler    行处理器
     * @param excel         是否是表格
     * @param <T>           类型
     * @return 数据读取器
     * @throws IOException 异常
     */
    public static <T> AbsDataReader<T, ?> readSync(MultipartFile multipartFile, AbsRowHandler<T> rowHandler, boolean excel) throws IOException, DataImportException, InterruptedException {
        String name = multipartFile.getOriginalFilename();
        if (StrUtil.isBlank(name)) {
            name = multipartFile.getName();
        }
        File importExcelTemp = ImportFileManager.createTempFile(FileUtil.getSuffix(name));
        FileUtil.writeFromStream(multipartFile.getInputStream(), importExcelTemp, false);
        if (excel) {
            return new ExcelDataReader<>(importExcelTemp, rowHandler).readSync();
        }
        return new WordTableDataReader<>(importExcelTemp, rowHandler).readSync();
    }

    /**
     * 读取表格信息信息
     *
     * @param multipartFile 文件
     * @param readClass     读取类
     * @param excel         是否是表格
     * @param <T>           类型
     * @return 数据读取器
     * @throws IOException 异常
     */
    public static <T> AbsDataReader<T, ?> readSync(MultipartFile multipartFile, Class<T> readClass, boolean excel) throws IOException, DataImportException, InterruptedException {
        String name = multipartFile.getOriginalFilename();
        if (StrUtil.isBlank(name)) {
            name = multipartFile.getName();
        }
        File importExcelTemp = ImportFileManager.createTempFile(FileUtil.getSuffix(name));
        FileUtil.writeFromStream(multipartFile.getInputStream(), importExcelTemp, false);
        RowHandlerDefault<T> rowHandlerDefault = new RowHandlerDefault<>(readClass);
        if (excel) {
            return new ExcelDataReader<>(importExcelTemp, rowHandlerDefault).readSync();
        }
        return new WordTableDataReader<>(importExcelTemp, rowHandlerDefault).readSync();
    }

    public static <T> RowHandlerDefault<T> createRowHandler(Class<T> readClass) {
        return new RowHandlerDefault<>(readClass);
    }

    public static boolean startsWith(String source, String... prefixList) {
        if (prefixList == null || prefixList.length < 1) {
            return false;
        }
        for (String prefix : prefixList) {
            if (source.startsWith(prefix)) {
                return true;
            }
        }
        return false;
    }

    public static boolean isEmpty(List<ReadCellInfo> dataList, int size) {
        if (dataList.isEmpty() || dataList.size() < size) {
            return true;
        }
        for (ReadCellInfo readCellInfo : dataList) {
            Object value = readCellInfo.getValue();
            if (value instanceof Collection) {
                value = ((Collection<?>) value).stream().filter(Objects::nonNull).map(StrUtil::toString).filter(StrUtil::isNotBlank).collect(Collectors.joining());
            }
            if (value != null && StrUtil.isNotBlank(StrUtil.toString(value))) {
                return false;
            }
        }
        return true;
    }
}
