package org.javaexcel.util;

import java.nio.ByteBuffer;
import java.util.UUID;

import org.apache.commons.codec.binary.Base64;

/*
 * File name : UUIDUtil.java
 * @Copyright : luoyoub@163.com 
 * Description : javaexcel 
 * Author : Robert
 * CreateTime : 2016年4月2日
 */
public final class UUIDUtil {
    /**
     * 获取唯一标识的UUID(32位)
     * 
     * @return
     */
    public static String getUUID() {
        String uid = UUID.randomUUID().toString();
        return uid.substring(0, 8) + uid.substring(9, 13) + uid.substring(14, 18) + uid.substring(19, 23)
                + uid.substring(24);
    }

    /**
     * 获取22位UUID
     * 
     * @return
     */
    public static String getShortUUID() {
        UUID uuid = UUID.randomUUID();
        return encode(uuid);
    }

    private static String encode(UUID uuid) {
        ByteBuffer buffer = ByteBuffer.wrap(new byte[16]);
        buffer.putLong(uuid.getMostSignificantBits());
        buffer.putLong(uuid.getLeastSignificantBits());
        return Base64.encodeBase64URLSafeString(buffer.array());
    }

    @SuppressWarnings("unused")
    private static UUID decode(String base64) {
        if (base64.length() != 22) {
            throw new IllegalArgumentException("Not a valid Base64 encoded UUID");
        }
        ByteBuffer buffer = ByteBuffer.wrap(Base64.decodeBase64(base64));
        if (buffer.capacity() != 16) {
            throw new IllegalArgumentException("Not a valid Base64 encoded UUID");
        }
        return new UUID(buffer.getLong(), buffer.getLong());
    }
}