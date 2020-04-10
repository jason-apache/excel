package com.jason.mapper;

import com.jason.base.CrudMapper;
import com.jason.entity.demo.Student;

import java.util.List;

/**
 * @Auther: Jason
 * @Date: 2019/12/28 23:11
 * @Description:
 */
public interface StudentMapper extends CrudMapper<Student> {

    List<Student> selectByClassesId(String cId);
}
