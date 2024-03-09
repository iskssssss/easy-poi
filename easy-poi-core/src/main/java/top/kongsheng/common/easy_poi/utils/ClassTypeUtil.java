package top.kongsheng.common.easy_poi.utils;

import cn.hutool.core.util.StrUtil;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Version 版权 Copyright(c)2024 KONG SHENG
 * @ClassName:
 * @Descripton:
 * @Author: 孔胜
 * @Date: 2021/10/09 09:54
 */
public final class ClassTypeUtil {
	private static final Map<String, ClassType> classTypeCacheMap = new ConcurrentHashMap<>();
	private static final String SIGNATURE_NAME = "signature";

	private static final String CONSTANT_1 = "/";
	private static final String CONSTANT_2 = ".";
	private static final String CONSTANT_3 = "<";
	private static final String CONSTANT_4 = ">";
	private static final String CONSTANT_5 = ";L";
	private static final String CONSTANT_6 = "<L";
	private static final String CONSTANT_7 = ">";

	private static final String CONSTANT_8 = "<T";
	private static final String CONSTANT_9 = ";T";

	private static final String CONSTANT_10 = ";";

	/**
	 * 获取signature字段的值
	 *
	 * @param obj 带获取对象
	 * @return signature字段的值
	 */
	public static String getSignature(Object obj) {
		try {
			final Field signatureField = obj.getClass().getDeclaredField(SIGNATURE_NAME);
			signatureField.setAccessible(true);
			final Object signatureValue = signatureField.get(obj);
			if (signatureValue instanceof String) {
				return ((String) signatureValue);
			}
			return null;
		} catch (Exception e) {
			return null;
		}
	}

	/**
	 * 获取obj的类型信息
	 *
	 * @param obj      待处理类
	 * @param startSub 是否先{@link ClassTypeUtil#matcher(String input, String pattern)}一遍
	 * @param isCache  是否启用缓存
	 * @return 类型信息
	 */
	public static ClassType getSignatureValue(Object obj, boolean startSub, boolean isCache) {
		try {
			String signature = ClassTypeUtil.getSignature(obj);
			if (signature == null) {
				Class<?> tType = null;
				if (obj instanceof Method) {
					final Method method = (Method) obj;
					tType = method.getReturnType();
				}
				if (obj instanceof Field) {
					final Field field = (Field) obj;
					tType = field.getType();
				}
				ClassType classType = new ClassType();
				classType.setTType(tType);
				return classType;
			}
			String signatureRes = signature;
			signatureRes = signatureRes.substring(signatureRes.indexOf(')') + 1).replace(CONSTANT_1, CONSTANT_2);
			ClassType classType = new ClassType();
			if (isCache) {
				// 从缓存获取已处理的类型。
				if (classTypeCacheMap.containsKey(signatureRes)) {
					return classTypeCacheMap.get(signatureRes);
				}
				// 将当前表达式的处理结果添加至缓存中。
				classTypeCacheMap.put(signatureRes, classType);
			}
			if (startSub) {
				signatureRes = matcher(signatureRes, "<(.*);>");
			} else {
				signatureRes = signatureRes.substring(0, signatureRes.length() - 1);
			}
			signatureRes = signatureRes.substring(1);
			signatureRes = signatureRes.replaceAll("(<L|<T)", CONSTANT_3);
			signatureRes = signatureRes.replaceAll("(;L|;T)", CONSTANT_10);
			signatureRes = signatureRes.replaceAll(";>", CONSTANT_4);
			// 处理类型表达式
			signatureHandler(classType, signatureRes);
			//System.out.println(classType.printStr());
			return classType;
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}

	public static ClassType signatureHandler(ClassType classType, String signature) throws ClassNotFoundException {
		if (isMapType(signature)) {
			// 处理非map类型
			String classStr;
			if (!signature.contains(CONSTANT_3)) {
				boolean forNameResult = forName(classType, signature);
				return classType;
			}
			int c2 = signature.indexOf(CONSTANT_3);
			classStr = signature.substring(0, c2);
			Class<?> aClass = Class.forName(classStr);
			classType.setTType(aClass);
			signature = matcher(signature, "<(.*)>");
			if (isMapType(signature)) {
				classType.setVType(signatureHandler(new ClassType(), signature));
				return classType;
			}
			signatureHandler(classType, signature);
			return classType;
		}
		// 处理Map  kv
		String[] signatureSplit = splitSignature_KV(signature);
		// 处理key类型
		String keyTypeStr = signatureSplit[0];
		ClassType keyClassType = setClassType(keyTypeStr);
		classType.setKType(keyClassType);
		// 处理value类型
		String valueTypeStr = signatureSplit[1];
		ClassType valueClassType = setClassType(valueTypeStr);
		classType.setVType(valueClassType);
		return classType;
	}

	public static ClassType setClassType(String typeStr) throws ClassNotFoundException {
		ClassType classType = new ClassType();
		if (typeStr.contains(CONSTANT_3)) {
			return signatureHandler(new ClassType(), typeStr);
		}
		boolean forNameResult = forName(classType, typeStr);
		return classType;
	}

	public static boolean forName(ClassType classType, String typeStr) {
		Class<?> keyType = null;
		try {
			keyType = Class.forName(typeStr);
		} catch (ClassNotFoundException ignored) { }
		if (keyType == null) {
			classType.setTTypeStr("*".equals(typeStr) ? "?" : typeStr);
			return false;
		}
		classType.setTType(keyType);
		return true;
	}

	public static String[] splitSignature_KV(String signature) {
		int length = signature.length();
		String[] ss = new String[2];
		int c1 = 0;
		for (int i = 0; i < length; i++) {
			char c = signature.charAt(i);
			if (c != '<' && c != '>' && c != ';') {
				continue;
			}
			if (c == '<') {
				c1++;
			} else if (c == '>') {
				c1--;
			} else if (c1 == 0) {
				return new String[]{signature.substring(0, i), signature.substring(i + 1)};
			}
		}
		return ss;
	}

	public static boolean isMapType(String signature) {
		int char1 = 0;
		int length = signature.length();
		for (int i = 0; i < length; i++) {
			char c = signature.charAt(i);
			if (c == ';') {
				char1++;
			}
			if (c == '>') {
				char1--;
			}
		}
		return char1 <= 0;
	}

	/**
	 * 匹配指定范围信息
	 *
	 * @param input 字符串
	 * @return 匹配信息
	 */
	public static String matcher(String input, String pattern) {
		if (StrUtil.isEmpty(input)) {
			return "";
		}
		Matcher matcher = Pattern.compile(pattern).matcher(input);
		if (!matcher.find()) {
			return "";
		}
		final String group1 = matcher.group(1);
		if (group1 == null) {
			return "";
		}
		return group1;
	}
}
