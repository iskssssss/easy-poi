package top.kongsheng.common.easy_poi.utils;

import lombok.Data;

/**
 * @Version 版权 Copyright(c)2024 KONG SHENG
 * @ClassName:
 * @Descripton:
 * @Author: 孔胜
 * @Date: 2021/10/08 11:32
 */
@Data
public class ClassType {

	public ClassType() {
	}

	public ClassType(Class<?> tType) {
		this.tType = tType;
	}

	public ClassType(String tTypeStr) {
		this.tTypeStr = tTypeStr;
	}

	/**
	 * 字段类型 T
	 */
	private Class<?> tType;

	/**
	 * 字段类型 字符串
	 */
	private String tTypeStr;

	/**
	 * 字段类型 K
	 */
	private ClassType kType;

	/**
	 * 字段类型 V
	 */
	private ClassType vType;

	public String printStr() {
		String simpleName = tTypeStr;
		if (tType != null) {
			simpleName = tType.getSimpleName();
		}
		if (kType != null && vType != null) {
			return simpleName + "<" + kType.printStr() + ", " + vType.printStr() + ">";
		}
		if (vType != null) {
			return simpleName + "<" + vType.printStr() + ">";
		}
		if (simpleName != null) {
			return simpleName;
		}
		throw new RuntimeException("参数有误");
	}
}
