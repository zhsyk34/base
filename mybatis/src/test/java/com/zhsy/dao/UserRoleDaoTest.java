package com.zhsy.dao;

import com.zhsy.entity.UserRole;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import javax.annotation.Resource;

/**
 * Created by Archimedes on 2016/5/22.
 */
@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration({"classpath:spring/spring-dao.xml"})
public class UserRoleDaoTest {

    @Resource
    private UserRoleDao userRoleDao;

    @Test
    public void save() throws Exception {
        UserRole userRole = new UserRole();
        userRole.setUserid(2);
        userRole.setRoleid(4);
        int result = userRoleDao.save(userRole);
        System.out.println(result);
    }

    @Test
    public void saveByAssociation() throws Exception {
        int result = userRoleDao.saveByAssociation(3, 4);
        System.out.println(result);
    }

    @Test
    public void delete() throws Exception {
        UserRole userRole = new UserRole();
        userRole.setUserid(3);
        userRole.setRoleid(4);
        int result = userRoleDao.delete(userRole);
        System.out.println(result);
    }

    @Test
    public void deleteByAssociation() throws Exception {
        int result = userRoleDao.deleteByAssociation(-1, -1);
        System.out.println(result);
    }
}
