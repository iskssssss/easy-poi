package top.kongsheng.common.easy_poi.value.handle;

import cn.hutool.core.util.StrUtil;
import top.kongsheng.common.easy_poi.model.ReadCellInfo;

import java.util.*;
import java.util.stream.Collectors;

/**
 * ValueHandlerByList
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/5 15:43
 */
public class ValueHandlerByList implements ValueHandler<List<String>> {

    @Override
    public List<String> to(String title, ReadCellInfo readCellInfo, Object... otherParams) {
        Object value = readCellInfo.getValue();
        List<String> resultList = null;
        if (value instanceof Collection) {
            Collection<?> valueList = (Collection<?>) value;
            resultList = this.replaceChar(valueList);
        }
        if (value instanceof String) {
            String valueStr = (String) value;
            resultList = this.handlerString(valueStr);
        }
        if (resultList == null) {
            return null;
        }
        return resultList.stream().distinct().collect(Collectors.toList());
    }

    @Override
    public void clear() {

    }

    /**
     *
     * @param valueStr
     * @return
     */
    private List<String> handlerString(String valueStr) {
        List<String> resultList = new LinkedList<>();
        int count = StrUtil.count(valueStr, "    ");
        if (count > 0) {
            valueStr = valueStr.replace(" ", "\n");
        }
        if (valueStr.contains("\n")) {
            List<String> collect = Arrays.stream(valueStr.split("\n")).collect(Collectors.toList());
            resultList.addAll(replaceChar(collect));
            return resultList;
        }
        boolean xx = valueStr.startsWith("★");
        int length = valueStr.length();
        StringBuilder sb = new StringBuilder();
        if ((xx ? (length - 1) : length) > 2 && valueStr.contains(" ")) {
            for (int i = 0; i < length; i++) {
                char c = valueStr.charAt(i);
                if (c == ' ') {
                    if ((sb.toString().replace("★", "")).length() > 1) {
                        resultList.add(sb.toString());
                        sb = new StringBuilder();
                    }
                    continue;
                }
                sb.append(c);
            }
            resultList.add(sb.toString());
            return replaceChar(resultList);
        }
        resultList.add(valueStr.replaceAll(" ", ""));
        return resultList;
    }

    private List<String> replaceChar(Collection<?> valueList) {
        List<String> resultListTemp = new ArrayList<>();
        for (Object value : valueList) {
            String valueStr = StrUtil.toString(value);
            if (StrUtil.isBlank(valueStr) || "null".equals(valueStr)) {
                continue;
            }
            List<String> handlerList = this.handlerString(valueStr);
            resultListTemp.addAll(handlerList);
        }
        int size = resultListTemp.size();
        if (size < 2) {
            return resultListTemp;
        }
        List<String> resultList = new ArrayList<>();
        for (int i = 0, step = 1; i < size; i = i + step, step = 1) {
            StringBuilder str = new StringBuilder();
            str.append(resultListTemp.get(i));
            boolean check = true;
            for (int j = i + 1; j < size; j++) {
                str.append(resultListTemp.get(j));
                step++;
                if (ValueHandlerByList.FAULT_TOLERANCE_SET.contains(str.toString())) {
                    resultList.add(str.toString());
                    check = false;
                    break;
                }
            }
            if (check) {
                step = 1;
                resultList.add(resultListTemp.get(i));
            }
        }
        return resultList;
    }

    public static final Set<String> FAULT_TOLERANCE_SET = new HashSet<>();
    static {
        FAULT_TOLERANCE_SET.add("自然资源和规划局");
        FAULT_TOLERANCE_SET.add("彭敏");
        FAULT_TOLERANCE_SET.add("李巍");
        FAULT_TOLERANCE_SET.add("农业农村局");
        FAULT_TOLERANCE_SET.add("大数据和金融发展中心");
        FAULT_TOLERANCE_SET.add("市生态环境局松阳分局");
        FAULT_TOLERANCE_SET.add("“五水共治”办");
        FAULT_TOLERANCE_SET.add("练斌");
        FAULT_TOLERANCE_SET.add("潘永水");
        FAULT_TOLERANCE_SET.add("潘毅");
        FAULT_TOLERANCE_SET.add("涂扬");
    }
}
