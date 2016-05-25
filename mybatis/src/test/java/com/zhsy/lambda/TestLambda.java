package com.zhsy.lambda;

import org.junit.Test;

import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Predicate;

/**
 * Created by Archimedes on 2016/5/18.
 */
/*
* *
一个 Lambda 表达式可以有零个或多个参数
参数的类型既可以明确声明，也可以根据上下文来推断。例如：(int a)与(a)效果相同
所有参数需包含在圆括号内，参数之间用逗号相隔。例如：(a, b) 或 (int a, int b) 或 (String a, int b, float c)
空圆括号代表参数集为空。例如：() -> 42
当只有一个参数，且其类型可推导时，圆括号（）可省略。例如：a -> return a*a
Lambda 表达式的主体可包含零条或多条语句
如果 Lambda 表达式的主体只有一条语句，花括号{}可省略。匿名函数的返回类型与该主体表达式一致
如果 Lambda 表达式的主体包含一条以上语句，则表达式必须包含在花括号{}中（形成代码块）。匿名函数的返回类型与代码块的返回类型一致，若没有返回则为空
* */
public class TestLambda {
    public static void evaluate(List<Integer> list, Predicate<Integer> predicate) {
        list.forEach((Integer i) -> {
            if (predicate.test(i)) {
                System.out.print(i + " ");
            }
        });
        System.out.println();
    }

    @Test
    public void test1() {

    }

    @Test
    public void testList() {
        List<Integer> list = Arrays.asList(1, 2, 3, 4, 5, 6, 7);

        list.forEach((n) -> {
            if (n > 5) {
                System.out.println(n);
            }
        });
    }

    @Test
    public void testList2() {
        List<Integer> list = Arrays.asList(1, 2, 3, 4, 5, 6, 7);
        for (Integer n : list) {
            System.out.println(n);
        }
        list.forEach((n) -> System.out.println(n));
        list.forEach(System.out::println);
    }

    @Test
    public void testList3() {
        List<Integer> list = Arrays.asList(1, 2, 3, 4, 5, 6, 7);
        evaluate(list, (Integer n) -> true);
        evaluate(list, (n) -> false);
        evaluate(list, (n) -> n % 2 == 0);
        evaluate(list, (n) -> n % 3 == 0);
        evaluate(list, (n) -> n > 3);
    }

    @Test
    public void testMap() {
        Map<String, Integer> map = new HashMap<>();
        map.put("A", 10);
        map.put("B", 20);
        map.put("C", 30);
        map.put("D", 40);
        map.put("E", 50);
        map.put("F", 60);

        map.forEach((k, v) -> System.out.println("Item : " + k + " Count : " + v));
    }

    @Test
    public void testRun() {
        for (int i = 1; i < 10; i++) {
            Runnable runnable = () -> {
                for (int k = 0; k < 10; k++) {
                    System.out.println(Thread.currentThread().getName() + ":" + k);
                }
            };
            new Thread(runnable, "thread" + i).start();
        }
    }
}
