package com.cat.audio;

import java.io.InputStream;

/**
 * Created by Administrator on 16-5-26.
 */
public class Util {

    public static String getPath(String name) {
        return Thread.currentThread().getContextClassLoader().getResource(name).getPath();
    }

    public static InputStream getStream(String name) {
        return Thread.currentThread().getContextClassLoader().getResourceAsStream(name);
    }
}
