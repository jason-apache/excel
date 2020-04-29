package com.jason.entity.demo;

import java.io.Serializable;

/**
 * @Auther: Jason
 * @Date: 2019/12/28 23:09
 * @Description:
 */
public class Student implements Serializable{

    private String id;
    private String name;
    private String cId;
    private Classes classes;

    public String getId() {
        return id;
    }

    public Student setId(String id) {
        this.id = id;
        return this;
    }

    public String getName() {
        return name;
    }

    public Student setName(String name) {
        this.name = name;
        return this;
    }

    public String getcId() {
        return cId;
    }

    public Student setcId(String cId) {
        this.cId = cId;
        return this;
    }

    public Classes getClasses() {
        return classes;
    }

    public Student setClasses(Classes classes) {
        this.classes = classes;
        return this;
    }

    @Override
    public String toString() {
        return "Student{" +
                "id='" + id + '\'' +
                ", name='" + name + '\'' +
                ", cId='" + cId + '\'' +
                ", classes=" + classes +
                '}';
    }
}
