package org.javaexcel.util;

import java.util.Map;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;

/*
 * File name   : JsonUtil.java
 * @Copyright  : luoyoub@163.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月2日
 */
public final class JsonUtil {
    private static Gson gson = new GsonBuilder().setDateFormat("yyyy-MM-dd HH:mm:ss").create();

    public static <T> T stringToBean(String json, Class<T> classOfT) {
        return gson.fromJson(json, classOfT);
    }

    public static String beanToString(Object object) {
        return gson.toJson(object);
    }

    public static void main(String[] args) {
        Object obj = new Object();
        String s = beanToString(obj);
        System.out.println(s);
        Map map = JsonUtil.stringToBean(s, Map.class);
        System.out.println(map);
    }
}