package test.excel;

import cn.hutool.core.io.FileUtil;
import lombok.Data;
import lombok.ToString;
import org.junit.Test;
import top.kongsheng.common.easy_poi.anno.PoiModel;
import top.kongsheng.common.easy_poi.anno.PoiModelField;
import top.kongsheng.common.easy_poi.enums.HorizontalAlignment;
import top.kongsheng.common.easy_poi.enums.VerticalAlignment;
import top.kongsheng.common.easy_poi.writer.DataWriteUtil;
import top.kongsheng.common.easy_poi.writer.abs.AbsWriterRowHandler;
import top.kongsheng.common.easy_poi.writer.model.RowHandleResult;

import java.io.*;
import java.util.LinkedList;
import java.util.List;

/**
 * 测试类
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2023/3/22 14:19
 */
public class TestMain {

    @Test
    public void test() throws IOException {
        boolean excel = false;
        String fileName = System.currentTimeMillis() + "." + (excel ? "xlsx" : "docx");
        File file = FileUtil.file("\\exports\\" + fileName);
        System.out.println(file.getAbsoluteFile());
        LinkedList<String> dataList = new LinkedList<>();
        for (int i = 0; i < 100; i++) {
            dataList.add("张" + i);
        }
        TestModelWriterRowHandler rowHandler = new TestModelWriterRowHandler(dataList);
        try (BufferedOutputStream outputStream = FileUtil.getOutputStream(file)) {
            DataWriteUtil.export( rowHandler, new LinkedList<>(),
                    true, new LinkedList<>(), outputStream, excel);
        }
    }

    public static class TestModelWriterRowHandler extends AbsWriterRowHandler<String, TestModel> {

        public TestModelWriterRowHandler(List<String> dataList) {
            super(dataList);
        }

        @Override
        public Class<TestModel> exportModelClass() {
            return TestModel.class;
        }

        @Override
        protected RowHandleResult<TestModel> to(long index, String lexicon) {
            TestModel result = new TestModel();
            result.setIndex(index);
            result.setName(lexicon);
            return new RowHandleResult<>(result);
        }
    }

    @Data
    @ToString
    @PoiModel(value = "测试", documentWidth = 11907, documentHeight = 23800)
    public static class TestModel {

        @PoiModelField(value = "序号", x = 0,
                writeConfig = @PoiModelField.WriteConfig(
                        widthRate = 0.1F,
                        vertical = top.kongsheng.common.easy_poi.enums.VerticalAlignment.CENTER,
                        horizontal = top.kongsheng.common.easy_poi.enums.HorizontalAlignment.CENTER)
        )
        private long index;

        @PoiModelField(value = "名称", x = 1,
                readConfig = @PoiModelField.ReadConfig(required = false),
                writeConfig = @PoiModelField.WriteConfig(
                        widthRate = 0.9F,
                        vertical = VerticalAlignment.CENTER,
                        horizontal = HorizontalAlignment.CENTER)
        )
        private String name;
    }
}
