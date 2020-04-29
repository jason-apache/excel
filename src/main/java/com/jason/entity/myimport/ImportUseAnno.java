package com.jason.entity.myimport;

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
@ExcelField(startRow = 2, sheetIndex = 1)
public class ImportUseAnno extends DataEntity<ImportUseAnno> {

    @ExcelField(title = "Integer")
    private Integer aInteger;

    private Long aLong;

    private Short aShort;

    private Byte aByte;

    private Double aDouble;

    private Float aFloat;

    private Boolean aBoolean;

    private Character aCharacter;

    private ImportUseAnno parent;

    private String template;

    @ExcelField(title = "parent",call = "name")
    private Student student;

    private Date date;

    public Integer getaInteger() {
        return aInteger;
    }

    public ImportUseAnno setaInteger(Integer aInteger) {
        this.aInteger = aInteger;
        return this;
    }

    public Long getaLong() {
        return aLong;
    }

    @ExcelField(title = "Long")
    public ImportUseAnno setaLong(Long aLong) {
        this.aLong = aLong;
        return this;
    }

    public Short getaShort() {
        return aShort;
    }

    @ExcelField(title = "Short")
    public ImportUseAnno setaShort(Short aShort) {
        this.aShort = aShort;
        return this;
    }

    public Byte getaByte() {
        return aByte;
    }

    @ExcelField(title = "Byte")
    public ImportUseAnno setaByte(Byte aByte) {
        this.aByte = aByte;
        return this;
    }

    public Double getaDouble() {
        return aDouble;
    }

    @ExcelField(title = "Double")
    public ImportUseAnno setaDouble(Double aDouble) {
        this.aDouble = aDouble;
        return this;
    }

    public Float getaFloat() {
        return aFloat;
    }

    @ExcelField(title = "Float")
    public ImportUseAnno setaFloat(Float aFloat) {
        this.aFloat = aFloat;
        return this;
    }

    public Boolean getaBoolean() {
        return aBoolean;
    }

    @ExcelField(title = "aBoolean")
    public ImportUseAnno setaBoolean(Boolean aBoolean) {
        this.aBoolean = aBoolean;
        return this;
    }

    public Character getaCharacter() {
        return aCharacter;
    }

    @ExcelField(title = "Char")
    public ImportUseAnno setaCharacter(Character aCharacter) {
        this.aCharacter = aCharacter;
        return this;
    }

    public ImportUseAnno getParent() {
        return parent;
    }

    @ExcelField(title = "parent", call = "student.classes" , callMethod = "setId",
            useTemplate = true,templateNameKey = "parent")
    public ImportUseAnno setParent(ImportUseAnno parent) {
        this.parent = parent;
        return this;
    }

    public String getTemplate() {
        return template;
    }

    @ExcelField(title = "字典数据",useTemplate = true)
    public ImportUseAnno setTemplate(String template) {
        this.template = template;
        return this;
    }

    public Date getDate() {
        return date;
    }

    public ImportUseAnno setDate(Date date) {
        this.date = date;
        return this;
    }

    @Override
    public String toString() {
        return "ImportUseAnno{" +
                "aInteger=" + aInteger +
                ", aLong=" + aLong +
                ", aShort=" + aShort +
                ", aByte=" + aByte +
                ", aDouble=" + aDouble +
                ", aFloat=" + aFloat +
                ", aBoolean=" + aBoolean +
                ", aCharacter=" + aCharacter +
                ", parent=" + parent +
                ", template='" + template + '\'' +
                ", date=" + date +
                '}';
    }
}
