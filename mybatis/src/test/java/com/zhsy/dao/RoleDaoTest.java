package com.zhsy.dao;

import com.zhsy.entity.Role;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

/**
 * Created by Archimedes on 2016/5/21.
 */
@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration({"classpath:spring/spring-dao.xml"})
public class RoleDaoTest {

    @Resource
    private RoleDao roleDao;

    @Test
    public void count() throws Exception {
        Map<String, Object> map = new HashMap<>();
        map.put("name like", "G%");//..
        int count = roleDao.count(map);

        System.out.println("count: " + count);
    }

    @Test
    public void deleteByEntity() throws Exception {
        Role role = roleDao.findById(14);
        int result = roleDao.deleteByEntity(role);
        System.out.println(result);
    }

    @Test
    public void deleteByEntitys() throws Exception {
        int[] ids = {13, 12, 11};
        List<Role> roles = new ArrayList<>();
        for (int id : ids) {
            Role role = roleDao.findById(id);
            if (role != null) {
                roles.add(role);
            }
        }
        roles.forEach(role -> System.out.println(role));
        int result = roleDao.deleteByEntitys(roles);
        System.out.println(result);
    }

    @Test
    public void deleteById() throws Exception {
        int result = roleDao.deleteById(1);
        System.out.println(result);
    }

    @Test
    public void deleteByIds() throws Exception {
        int[] ids = {13, 15, 17, 18};
        int result = roleDao.deleteByIds(ids);
        System.out.println(result);
    }

    @Test
    public void findAll() throws Exception {
        List<Role> roles = (List<Role>) roleDao.findAll();
        roles.forEach(role -> System.out.println(role));
    }

    @Test
    public void findById() throws Exception {
        System.out.println(roleDao.findById(1));
    }

    @Test
    public void findInterval() throws Exception {
        List<Role> roles = (List<Role>) roleDao.findInterval(1, 4);
        roles.forEach(role -> System.out.println(role));
    }

    @Test
    public void save() throws Exception {
        Role role = new Role();
        role.setName("菜鸟角色");
        int result = roleDao.save(role);
        System.out.println(result);
    }

    @Test
    public void saves() throws Exception {
        List<Role> roles = new ArrayList<>();
        Role role;
        for (int i = 1; i < 5; i++) {
            role = new Role();
            role.setName("普通玩家" + i);
            roles.add(role);
        }

        int result = roleDao.saves(roles);
        System.out.println(result);
    }

    @Test
    public void update() throws Exception {
        Role role = roleDao.findById(7);
        role.setName("简单变更玩家");
        int result = roleDao.update(role);
        System.out.println(result);
    }

    @Test
    public void updates() throws Exception {
        List<Role> roleList = new ArrayList<>();
        for (int i = 11; i < 15; i++) {
            Role role = roleDao.findById(i);
            if (role == null) {
                continue;
            }
            role.setName("批量玩家" + i);
            roleList.add(role);
        }
        roleList.forEach((role) -> System.out.println(role));

        int result = roleDao.updates(roleList);
        System.out.println(result);
    }
}
