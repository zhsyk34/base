package com.zhsy.dao;

import com.zhsy.entity.UserRole;

import org.apache.ibatis.annotations.Param;

/**
 * Created by Archimedes on 2016/5/21.
 */
public interface UserRoleDao {

    int save(UserRole userRole);

    int saveByAssociation(int userid, int roleid);

    int delete(UserRole userRole);

    int deleteByAssociation(@Param("userid") int userid, @Param("roleid") int roleid);
}
