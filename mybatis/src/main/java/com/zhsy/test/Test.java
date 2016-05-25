package com.zhsy.test;

import com.zhsy.entity.User;

import org.apache.ibatis.io.Resources;
import org.apache.ibatis.session.SqlSession;
import org.apache.ibatis.session.SqlSessionFactory;
import org.apache.ibatis.session.SqlSessionFactoryBuilder;

import java.io.IOException;
import java.io.InputStream;

/**
 * Created by Archimedes on 2016/5/15.
 */
public class Test {

    public static void main(String[] args) throws IOException {

    }

    public void test1() throws IOException {
        String resource = "mybatis.xml";
        InputStream inputStream = Resources.getResourceAsStream(resource);
        SqlSessionFactory sqlSessionFactory = new SqlSessionFactoryBuilder().build(inputStream);

        System.out.println(sqlSessionFactory);

        SqlSession sqlSession = sqlSessionFactory.openSession();
        System.out.println(sqlSession);

        User user = null;
        //generate method:1
//        UserDao userDao = sqlSession.getMapper(UserDao.class);
//         user = userDao.find(1);

        //method:2
        user = sqlSession.selectOne("com.zhsy.dao.UserDao.find", 1);
        System.out.println(user);
    }
}