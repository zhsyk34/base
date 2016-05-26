package com.cat.demo;

import com.google.common.base.Charsets;
import com.google.common.base.Joiner;

/**
 * Created by Administrator on 16-5-26.
 */
public class TestGuava {

    public static void main(String[] args) {
        test2();
    }

    public static void test2() {
        String string = "Hello world!";
        byte[] bytes = string.getBytes(Charsets.UTF_8);
        for (int i = 0, len = bytes.length; i < len; i++) {
            System.out.print((char) bytes[i] + " ");
        }
    }

    public static void test1() {
        Joiner joiner = Joiner.on("; ").skipNulls();
        String result = joiner.join("Harry", null, "Ron", "Hermione");

        System.out.println(result);
    }
}
