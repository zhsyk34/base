package com.zhsy.excel;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;

/**
 * Created by Administrator on 16-5-25.
 */
public class CreateData {

    public static void main(String[] args) throws JsonProcessingException {

    }

    public static List<User> test2() throws JsonProcessingException {
        List<User> list = new ArrayList<>();
        for (int i = 1; i < 10; i++) {
            User user = new User(i, "name" + i, new Random().nextInt(100), "man", new Random().nextFloat());
            list.add(user);
        }

        ObjectMapper mapper = new ObjectMapper();
        String s = mapper.writeValueAsString(list);
        System.out.println(s);

        return list;
    }

    public static void test1() throws JsonProcessingException {
        Map<String, Object> map = new HashMap<>();
        map.put("age", 22);
        map.put("name", "jack");
        map.put("sex", "man");
        map.put("price", 32.5f);

        ObjectMapper mapper = new ObjectMapper();
        String s = mapper.writeValueAsString(map);
        System.out.println(s);
    }
}
