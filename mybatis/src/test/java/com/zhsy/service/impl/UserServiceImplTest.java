package com.zhsy.service.impl;

import com.zhsy.entity.User;
import com.zhsy.service.UserService;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.slf4j.LoggerFactory;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import java.util.List;

import javax.annotation.Resource;

/**
 * Created by Archimedes on 2016/5/22.
 */
@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration({"classpath:spring/spring-dao.xml", "classpath:spring/spring-service.xml"})
public class UserServiceImplTest {

    private final org.slf4j.Logger logger = LoggerFactory.getLogger(this.getClass());
    @Resource
    private UserService userService;

    @Test
    public void save() throws Exception {
        User user = new User();
        user.setAge(100);
        user.setName("张三丰");
        userService.save(user);
    }

    @Test
    public void saves() throws Exception {

    }

    @Test
    public void deleteById() throws Exception {
        userService.deleteById(3);
    }

    @Test
    public void deleteByEntity() throws Exception {

    }

    @Test
    public void deleteByIds() throws Exception {

    }

    @Test
    public void deleteByEntitys() throws Exception {

    }

    @Test
    public void update() throws Exception {

    }

    @Test
    public void updates() throws Exception {

    }

    @Test
    public void findById() throws Exception {

    }

    @Test
    public void findAll() throws Exception {
        List<User> users = userService.findAll();
        logger.info("users = {}", users);
    }

    @Test
    public void findByPage() throws Exception {

    }

    @Test
    public void count() throws Exception {

    }
}
