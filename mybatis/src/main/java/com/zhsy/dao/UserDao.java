package com.zhsy.dao;

import com.zhsy.entity.Role;
import com.zhsy.entity.User;

import java.util.Collection;

/**
 * Created by Archimedes on 2016/5/15.
 */
public interface UserDao extends CommonDao<User> {

    Collection<Role> findRoles(int id);
}
