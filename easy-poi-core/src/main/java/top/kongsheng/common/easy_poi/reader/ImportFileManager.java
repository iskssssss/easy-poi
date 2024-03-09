package top.kongsheng.common.easy_poi.reader;

import cn.hutool.cache.CacheUtil;
import cn.hutool.cache.impl.TimedCache;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.IdUtil;

import java.io.File;
import java.util.Optional;

/**
 * 导入文件管理类
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/5/6 14:33
 */
public class ImportFileManager {
    public final static String PREFIX = "EXCEL_IMPORT_FILE_";

    /**
     * 导入文件
     */
    private static final TimedCache<String, File> IMPORT_FILE_MAP = CacheUtil.newTimedCache((30 * 60) * 1000);

    static {
        String tmpDirPath = FileUtil.getTmpDirPath();
        File tmpDir = new File(tmpDirPath);
        Optional.ofNullable(tmpDir.listFiles()).ifPresent(files -> {
            for (File file : files) {
                String fileName = file.getName();
                if (fileName.startsWith(ImportFileManager.PREFIX)) {
                    FileUtil.del(file);
                    System.out.println("删除数据导入临时文件：" + file.getName());
                }
            }
        });
        IMPORT_FILE_MAP.setListener((key, file) -> {
            FileUtil.del(file);
            System.out.println("删除数据导入临时文件：" + file.getName());
        });
    }

    public static File find(String fileId) {
        return IMPORT_FILE_MAP.get(fileId);
    }

    public static String put(File file) {
        String fileId = IdUtil.fastSimpleUUID() + System.currentTimeMillis();
        IMPORT_FILE_MAP.put(fileId, file);
        return fileId;
    }

    public static void del(String fileId) {
        try {
            File file = IMPORT_FILE_MAP.get(fileId);
            if (file.exists()) {
                FileUtil.del(file);
            }
        } catch (Exception e) {
            System.err.println("导入文件删除失败");
        }
    }

    public static File createTempFile(String suffix) {
        return FileUtil.createTempFile(PREFIX, "." + suffix, null, true);
    }
}
