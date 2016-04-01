package org.javaexcel.util;

import com.google.gson.Gson;

/**
 * @author cuikexiang
 *
 * @time 2015年10月30日 上午10:37:03
 */
public final class JsonUtil {
    private static Gson gson = new Gson();

    public static <T> T stringToBean(String json, Class<T> classOfT) {
        return gson.fromJson(json, classOfT);
    }

    public static String beanToString(Object object) {
        return gson.toJson(object);
    }
}