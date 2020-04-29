package com.jason.entity.demo;

import com.jason.base.BaseEntity;
import com.jason.base.DataEntity;

import java.util.List;

/**
 * @Auther: Jason
 * @Date: 2019/12/28 23:10
 * @Description:
 */
public class Classes extends DataEntity<Classes> {

    private String name;
    private List<Student> studentList;

    public Classes(){
        super();
    }

    public Classes(String id){
        super(id);
    }

    public String getName() {
        return name;
    }

    public Classes setName(String name) {
        this.name = name;
        return this;
    }

    @Override
    public BaseEntity<Classes> setId(String id) {
        return super.setId(id);
    }

    public List<Student> getStudentList() {
        return studentList;
    }

    public Classes setStudentList(List<Student> studentList) {
        this.studentList = studentList;
        return this;
    }

    @Override
    public String toString() {
        return "Classes{" +
                "id='" + id + '\'' +
                ", name='" + name + '\'' +
                ", studentList=" + studentList +
                '}';
    }
}
