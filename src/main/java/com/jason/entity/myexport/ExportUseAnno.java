package com.jason.entity.myexport;

import com.jason.anno.ExcelField;
import com.jason.base.DataEntity;
import com.jason.entity.demo.Student;

import java.util.Date;

/**
 * @Auther: Jason
 * @Date: 2020/3/31 14:34
 * 使用注解
 * @Description:
 */
public class ExportUseAnno extends DataEntity<ExportUseAnno> {

    @ExcelField(title = "Integer",isImport = false,sort = 100000000)
    private Integer aInteger;

    @ExcelField(title = "Long",isImport = false,sort = 1)
    private Long aLong;

    @ExcelField(title = "Short",isImport = false,sort = 2)
    private Short aShort;

    @ExcelField(title = "Byte",isImport = false,sort = 3,useTemplate = true,templateNameKey = "第二个")
    private Byte aByte;

    @ExcelField(title = "Double",isImport = false,sort = 4)
    private Double aDouble;

    @ExcelField(title = "Float",isImport = false,sort = 5)
    private Float aFloat;

    @ExcelField(title = "aBoolean",isImport = false,sort = 6)
    private Boolean aBoolean;

    @ExcelField(title = "Char",isImport = false,sort = 7)
    private Character aCharacter;

    @ExcelField(title = "parent",call = "student.classes",callMethod = "getName",
            isImport = false,sort = 8)
    private ExportUseAnno parent;

    private String template;

    private Student student;

    @ExcelField(title = "date",isImport = false,sort = 10)
    private Date date;

    public Integer getaInteger() {
        return aInteger;
    }

    public ExportUseAnno setaInteger(Integer aInteger) {
        this.aInteger = aInteger;
        return this;
    }

    public Long getaLong() {
        return aLong;
    }

    public ExportUseAnno setaLong(Long aLong) {
        this.aLong = aLong;
        return this;
    }

    public Short getaShort() {
        return aShort;
    }

    public ExportUseAnno setaShort(Short aShort) {
        this.aShort = aShort;
        return this;
    }

    public Byte getaByte() {
        return aByte;
    }

    public ExportUseAnno setaByte(Byte aByte) {
        this.aByte = aByte;
        return this;
    }

    public Double getaDouble() {
        return aDouble;
    }

    public ExportUseAnno setaDouble(Double aDouble) {
        this.aDouble = aDouble;
        return this;
    }

    public Float getaFloat() {
        return aFloat;
    }

    public ExportUseAnno setaFloat(Float aFloat) {
        this.aFloat = aFloat;
        return this;
    }

    public Boolean getaBoolean() {
        return aBoolean;
    }

    public ExportUseAnno setaBoolean(Boolean aBoolean) {
        this.aBoolean = aBoolean;
        return this;
    }

    public Character getaCharacter() {
        return aCharacter;
    }

    public ExportUseAnno setaCharacter(Character aCharacter) {
        this.aCharacter = aCharacter;
        return this;
    }

    public ExportUseAnno getParent() {
        return parent;
    }

    public ExportUseAnno setParent(ExportUseAnno parent) {
        this.parent = parent;
        return this;
    }

    @ExcelField(title = "字典数据",isImport = false,sort = 9,useTemplate = true,templateNameKey = "default")
    public String getTemplate() {
        return template;
    }

    public ExportUseAnno setTemplate(String template) {
        this.template = template;
        return this;
    }

    public Date getDate() {
        return date;
    }

    public ExportUseAnno setDate(Date date) {
        this.date = date;
        return this;
    }

    public Student getStudent() {
        return student;
    }

    public ExportUseAnno setStudent(Student student) {
        this.student = student;
        return this;
    }
}
