package com.jason.entity;

import com.jason.anno.ExcelField;
import com.jason.base.DataEntity;
import com.jason.entity.demo.Student;

import java.util.Date;

/**
 * @author: Jason
 * @Date: 2020/4/29 11:17
 * @Description: 通用导入导出excel实体类
 */
@ExcelField(startRow = 2, sheetIndex = 1)
public class GeneralExcel extends DataEntity<GeneralExcel> {

    @ExcelField(title = "Integer",sort = 99999)
    private Integer aInteger;

    @ExcelField(title = "Long",sort = 1)
    private Long aLong;

    @ExcelField(title = "Short",sort = 2)
    private Short aShort;

    @ExcelField(title = "Byte",sort = 3)
    private Byte aByte;

    @ExcelField(title = "Double",sort = 4)
    private Double aDouble;

    @ExcelField(title = "Float",sort = 5)
    private Float aFloat;

    @ExcelField(title = "Boolean",sort = 6)
    private Boolean aBoolean;

    @ExcelField(title = "Character",sort = 7)
    private Character aCharacter;

    @ExcelField(title = "parent",sort = 8,call = "student.classes.name")
    private GeneralExcel parent;

    @ExcelField(title = "Date",sort = 9)
    private Date date;

    @ExcelField(title = "字典列1",sort = 10,useTemplate = true)
    private String dictData1;

    @ExcelField(title = "字典列2",sort = 11,useTemplate = true,templateNameKey = "第二个")
    private String dictData2;

    @ExcelField(title = "学生姓名",sort = 12,call = "name")
    private Student student;

    public Integer getaInteger() {
        return aInteger;
    }

    public GeneralExcel setaInteger(Integer aInteger) {
        this.aInteger = aInteger;
        return this;
    }

    public Long getaLong() {
        return aLong;
    }

    public GeneralExcel setaLong(Long aLong) {
        this.aLong = aLong;
        return this;
    }

    public Short getaShort() {
        return aShort;
    }

    public GeneralExcel setaShort(Short aShort) {
        this.aShort = aShort;
        return this;
    }

    public Byte getaByte() {
        return aByte;
    }

    public GeneralExcel setaByte(Byte aByte) {
        this.aByte = aByte;
        return this;
    }

    public Double getaDouble() {
        return aDouble;
    }

    public GeneralExcel setaDouble(Double aDouble) {
        this.aDouble = aDouble;
        return this;
    }

    public Float getaFloat() {
        return aFloat;
    }

    public GeneralExcel setaFloat(Float aFloat) {
        this.aFloat = aFloat;
        return this;
    }

    public Boolean getaBoolean() {
        return aBoolean;
    }

    public GeneralExcel setaBoolean(Boolean aBoolean) {
        this.aBoolean = aBoolean;
        return this;
    }

    public Character getaCharacter() {
        return aCharacter;
    }

    public GeneralExcel setaCharacter(Character aCharacter) {
        this.aCharacter = aCharacter;
        return this;
    }

    public GeneralExcel getParent() {
        return parent;
    }

    @ExcelField(title = "parent",isImport = true,
            call = "student.classes",callMethod = "setName",callClass = String.class)
    public GeneralExcel setParent(GeneralExcel parent) {
        this.parent = parent;
        return this;
    }

    public Date getDate() {
        return date;
    }

    public GeneralExcel setDate(Date date) {
        this.date = date;
        return this;
    }

    public String getDictData1() {
        return dictData1;
    }

    public GeneralExcel setDictData1(String dictData1) {
        this.dictData1 = dictData1;
        return this;
    }

    public String getDictData2() {
        return dictData2;
    }

    public GeneralExcel setDictData2(String dictData2) {
        this.dictData2 = dictData2;
        return this;
    }

    public Student getStudent() {
        return student;
    }

    public GeneralExcel setStudent(Student student) {
        this.student = student;
        return this;
    }
}
