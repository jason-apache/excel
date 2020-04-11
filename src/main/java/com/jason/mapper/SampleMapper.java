package com.jason.mapper;

import org.apache.ibatis.annotations.Param;

import java.util.LinkedHashMap;
import java.util.List;

/**
 * @Auther: Jason
 * @Date: 2020/4/10 20:49
 * 通用执行sql语句mapper
 * @Description:
 */
public interface SampleMapper {

    void DML(@Param("sql") String sql);

    String getSingleColumnString(@Param("sql") String sql);

    List<String> selectSingleColumnStringList(@Param("sql") String sql);

    List<LinkedHashMap<String,Object>> selectObject(@Param("sql") String sql);

    LinkedHashMap<String,Object> selectSingleObject(@Param("sql") String sql);
}
