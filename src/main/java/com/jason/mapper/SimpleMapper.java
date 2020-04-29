package com.jason.mapper;

import org.apache.ibatis.annotations.Param;

import java.util.LinkedHashMap;
import java.util.List;

/**
 * @author: Jason
 * @Date: 2020/4/10 20:49
 * 通用执行sql语句mapper
 * @Description:
 */
public interface SimpleMapper {

    /**
     * 增删改操作
     * @param sql
     */
    void DML(@Param("sql") String sql);

    /**
     * 获取单个字符串
     * @param sql
     * @return 数据库中的单个格
     */
    String getSingleColumnString(@Param("sql") String sql);

    /**
     * 获取单列
     * @param sql
     * @return 数据库中的单列
     */
    List<String> selectSingleColumnStringList(@Param("sql") String sql);

    /**
     * 获取单条记录，转换成map键值对的形式
     * @param sql
     * @return 数据库中的某一行
     */
    LinkedHashMap<String,Object> selectSingleObject(@Param("sql") String sql);

    /**
     * 获取列表记录，转换成map键值对的形式
     * @param sql
     * @return 数据库中的行
     */
    List<LinkedHashMap<String,Object>> selectObject(@Param("sql") String sql);
}
