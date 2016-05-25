package com.zhsy.service;

import com.zhsy.entity.Role;
import com.zhsy.entity.User;

import java.util.List;

/**
 * Created by Archimedes on 2016/5/22.
 */
public interface UserService {

    int save(User user);

    int saves(List<User> users);

    int deleteById(int id);

    int deleteByEntity(User user);

    int deleteByIds(int[] ids);

    int deleteByEntitys(List<User> users);

    int update(User user);

    int updates(List<User> users);

    User findById(int id);

    List<User> findAll();

    List<User> findByPage(int pageNo, int pageSize);

    int count(int pageNo, int pageSize);

    List<Role> findRoles(int id);
}
