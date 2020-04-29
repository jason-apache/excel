package com.jason.mapper;

import com.jason.base.CrudMapper;
import com.jason.entity.demo.Student;

import java.util.List;

/**
 * @author: Jason
 * @Date: 2019/12/28 23:11
 * @Description:
 */
public interface StudentMapper extends CrudMapper<Student> {

    /**
    * @author Jason
    * @date 2020/4/29 10:38
    * @params [cId]
    * 根据classId查询出所有学生
    * @return java.util.List<com.jason.entity.demo.Student>
    */
    List<Student> selectByClassesId(String cId);
}
