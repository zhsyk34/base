package com.zhsy.dao;

import java.util.Collection;
import java.util.Map;

/**
 * Created by Archimedes on 2016/5/21.
 */
public interface CommonDao<Entity> {

    int save(Entity entity);

    int saves(Collection<Entity> entities);

    int deleteById(int id);

    int deleteByEntity(Entity entity);

    int deleteByIds(int[] ids);

    int deleteByEntitys(Collection<Entity> entities);

    int update(Entity entity);

    int updates(Collection<Entity> entities);

    Entity findById(int id);

    Collection<Entity> findAll();

    Collection<Entity> findInterval(int offset, int limit);

    int count(Map<String, Object> map);
}
