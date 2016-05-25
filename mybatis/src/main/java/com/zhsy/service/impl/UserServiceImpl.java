package com.zhsy.service.impl;

import com.zhsy.dao.UserDao;
import com.zhsy.dao.UserRoleDao;
import com.zhsy.entity.Role;
import com.zhsy.entity.User;
import com.zhsy.service.UserService;

import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

/**
 * Created by Archimedes on 2016/5/22.
 */
@Service
public class UserServiceImpl implements UserService {

    @Resource
    private UserDao userDao;
    @Resource
    private UserRoleDao userRoleDao;

    @Override
    public int save(User user) {
        return userDao.save(user);
    }

    @Override
    public int saves(List<User> users) {
        return userDao.saves(users);
    }

    @Transactional
    @Override
    public int deleteById(int id) {
        userRoleDao.deleteByAssociation(id, 0);
        return userDao.deleteById(id);
    }

    @Override
    public int deleteByEntity(User user) {
        return userDao.deleteByEntity(user);
    }

    @Override
    public int deleteByIds(int[] ids) {
        return userDao.deleteByIds(ids);
    }

    @Override
    public int deleteByEntitys(List<User> users) {
        return userDao.deleteByEntitys(users);
    }

    @Override
    public int update(User user) {
        return userDao.update(user);
    }

    @Override
    public int updates(List<User> users) {
        return userDao.updates(users);
    }

    @Override
    public User findById(int id) {
        return userDao.findById(id);
    }

    @Override
    public List<User> findAll() {
        return (List<User>) userDao.findAll();
    }

    @Override
    public List<User> findByPage(int pageNo, int pageSize) {
        return pageNo > 0 && pageSize > 0 ? (List<User>) userDao.findInterval((pageNo - 1) * pageSize, pageNo * pageSize) : null;
    }

    @Override
    public int count(int pageNo, int pageSize) {
        Map<String, Object> map = new HashMap<>();
        map.put("limit", (pageNo - 1) * pageSize);
        map.put(" ", (pageNo) * pageSize);
        return userDao.count(map);
    }

    @Override
    public List<Role> findRoles(int id) {
        return (List<Role>) userDao.findRoles(id);
    }
}
