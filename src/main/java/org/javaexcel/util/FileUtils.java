package org.javaexcel.util;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/*
 * File name   : FileUtils.java
 * @Copyright  : luoyoub@163.com
 * Description : javaexcel
 * Author      : Robert
 * CreateTime  : 2016年4月7日
 */
public class FileUtils {
    @SuppressWarnings("unchecked")
    public static void substitute(String zipfile, String tmpfile,
            String entry, OutputStream out) throws IOException {
        ZipFile zip = new ZipFile(zipfile);
        ZipOutputStream zos = new ZipOutputStream(out);
        Enumeration<ZipEntry> en = (Enumeration<ZipEntry>) zip.entries();
        while (en.hasMoreElements()) {
            ZipEntry ze = en.nextElement();
            if (!ze.getName().equals(entry)) {
                zos.putNextEntry(new ZipEntry(ze.getName()));
                InputStream is = zip.getInputStream(ze);
                copyStream(is, zos);
                is.close();
            }
        }
        zos.putNextEntry(new ZipEntry(entry));
        InputStream is = new FileInputStream(tmpfile);
        copyStream(is, zos);
        zip.close();
        is.close();
        zos.close();
    }

    private static void copyStream(InputStream in, OutputStream out)
            throws IOException {
        byte[] chunk = new byte[1024];
        int count;
        while ((count = in.read(chunk)) >= 0) {
            out.write(chunk, 0, count);
        }
    }
}