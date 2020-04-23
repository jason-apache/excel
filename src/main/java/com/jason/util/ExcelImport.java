package com.jason.util;

import com.jason.anno.ExcelField;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 *
 * @Author jason
 * @createTime 2019年12月18日 21:09
 * @Description
 */
public class ExcelImport<T> {
    /**
     * 解析起始行
     */
    private int startRow;
    /**
     * 解析工作簿
     */
    private int startSheet;
    /**
     *工作簿名称
     */
    private String sheetName;
    /**
     *实体类型
     */
    private final Class<T> clazz;
    /**
     *输入流
     */
    private final InputStream is;
    /**
     *字节数组，保存流
     */
    private byte[] bytes;

    private Sheet sheet;
    /**
     *是否初始化
     */
    private boolean initialized;
    /**
     *自动根据字段名称映射
     */
    private boolean autoMappingByFieldName = true;
    /**
     *注解
     */
    private List<ExcelField> annotationList;
    /**
     *注解映射关系 ExcelField -> field or method
     */
    private Map<ExcelField,Object> annotationMapping;
    /**
     *title映射关系
     */
    private Map<String,Integer> titleMapping;
    /**
     *不声明方法式设值时，默认以字段名映射excel
     */
    private Set<Field> fieldsSet;
    /**
     *模板格式
     */
    private Map<String,Map<String,String>> template;

    public ExcelImport(InputStream is,Class<T> clazz){
        ExcelField field = clazz.getAnnotation(ExcelField.class);
        //根据注解中的属性设初值
        if(null != field){
            startRow = field.startRow() > 0 ? field.startRow() - 1 : 0;
            startSheet = field.sheetIndex() > 0 ? field.sheetIndex() - 1 : 0;
            sheetName = "".equals(field.sheetName().trim()) ? null : field.sheetName();
        }
        this.clazz = clazz;
        this.is = is;
        this.initMethods();
    }

    /**
     * @author Jason
     * @date 2020/4/22 10:18
     * @params []
     * 初始化工作薄
     * @return void
     */
    private void init() throws IOException {
        try {
            bytes = new byte[is.available()];
            is.read(bytes);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(new ByteArrayInputStream(bytes));
            //初始化
            if(StringUtil.isNotBlank(sheetName)){
                sheet = xssfWorkbook.getSheet(sheetName);
            }else{
                sheet = xssfWorkbook.getSheetAt(startSheet);
            }
        }catch (Exception e){
            System.out.println("格式不匹配");
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(new ByteArrayInputStream(bytes));
            //初始化
            if(StringUtil.isNotBlank(sheetName)){
                sheet = hssfWorkbook.getSheet(sheetName);
            }else{
                sheet = hssfWorkbook.getSheetAt(startSheet);
            }
        }

        //取excel首行列名
        Row firstRow = sheet.getRow(startRow);
        titleMapping = new HashMap<>(sheet.getLastRowNum());
        //取出excel列的位置index，放入title映射
        for(int i=0;i<firstRow.getLastCellNum();i++){
            String data = firstRow.getCell(i) == null ? "" : firstRow.getCell(i).toString();
            titleMapping.put(data,i);
        }
        this.initialized = true;
    }

    /**
     * @author Jason
     * @date 2020/4/22 10:19
     * @params []
     * 初始化方法
     * @return void
     */
    private void initMethods(){

        Field[] fields = clazz.getDeclaredFields();
        Method[] methods = clazz.getDeclaredMethods();
        annotationList = new ArrayList<>(fields.length + methods.length);
        annotationMapping = new HashMap<>(fields.length + methods.length);
        //自动根据字段名映射
        for (Method method : methods) {
            ExcelField excelField = method.getAnnotation(ExcelField.class);
            if (null != excelField && excelField.isImport() && StringUtil.isNotBlank(excelField.title())) {
                annotationList.add(excelField);
                annotationMapping.put(excelField, method);
            }
        }
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (null != excelField && excelField.isImport() && StringUtil.isNotBlank(excelField.title())) {
                annotationList.add(excelField);
                annotationMapping.put(excelField, field);
            } else {
                if (autoMappingByFieldName) {
                    if (null == fieldsSet) {
                        fieldsSet = new HashSet<>();
                    }
                    fieldsSet.add(field);
                }
            }
        }
    }

    /**
     * @author Jason
     * @date 2020/3/31 13:21
     * @params [file, startRow, startSheet, collection]
     * 转为Java对象 返回错误信息
     * @return String
     */
    public String getObjects(Collection<T> collection) throws IOException {

        if(!this.initialized){
            this.init();
        }
        StringBuilder errorMsg = new StringBuilder();
        for(int i=this.startRow+1;i<this.sheet.getLastRowNum()+1;i++){
            try {
                T o = this.getObject(this.sheet.getRow(i));
                if(null != o){
                    collection.add(o);
                }
            }catch (Exception e){
                errorMsg.append("错误信息：第").append(i).append("行，").append(e.getMessage()).append("\r\n");
            }
        }
        return errorMsg.toString();
    }

    /**
     * @author Jason
     * @date 2020/3/30 17:18
     * @params [row]
     * 解析excel
     * @return T
     */
    public T getObject(Row row) throws IllegalAccessException, InstantiationException,
            NoSuchMethodException, InvocationTargetException, IOException, ParseException {
        if(null == row || row.getLastCellNum() == 0){
            return null;
        }
        if(!this.initialized){
            this.init();
        }
        T t = clazz.newInstance();
        //根据参数位置映射，开始解析excel
        for (ExcelField excelField : annotationList) {
            if (null != excelField) {
                Object o = annotationMapping.get(excelField);
                Cell cell = null;
                //如果使用了position属性
                if (excelField.position() != -1) {
                    cell = row.getCell(excelField.position());
                } else {
                    Integer index = titleMapping.get(excelField.title());
                    if (null != index) {
                        cell = row.getCell(titleMapping.get(excelField.title()));
                    } else {
                        continue;
                    }
                }

                this.setValue(o,excelField,cell,t);
            }
        }
        //是否自动根据参数名映射 默认开启
        if(autoMappingByFieldName){
            for(Field f : fieldsSet){
                f.setAccessible(true);
                Integer index = titleMapping.get(f.getName());
                if(null != index){
                    if(null == f.get(t)){
                        Cell cell = row.getCell(index);
                        this.setValue(f,cell,t);
                    }
                }
            }
        }
        return t;
    }

    /**
     * @author Jason
     * @date 2020/4/23 14:17
     * @params [o, excelField, cell, t]
     * @return void
     * 设值，过滤模板格式
     */
    private void setValue(Object o,ExcelField excelField, Cell cell, T t)
            throws IllegalAccessException, ParseException, InvocationTargetException, NoSuchMethodException, InstantiationException {
        Object val = null;
        if(excelField.useTemplate()){
            Map<String, String> map = template.get(excelField.templateNameKey());
            if(null != map){
                val = map.get(cell.toString());
            }
        }

        if(o instanceof Method){
            this.setValue((Method) o,excelField,cell,val,t);
        }else if(o instanceof Field){
            this.setValue((Field) o,excelField,cell,val,t);
        }
    }

    /**
     * @author Jason
     * @date 2020/4/23 14:17
     * @params [field, excelField, cell, val, t]
     * @return void
     * 设值
     */
    private void setValue(Field field,ExcelField excelField,Cell cell,Object val,T t) throws ParseException, IllegalAccessException {
        field.setAccessible(true);
        if(excelField.useTemplate()){
            field.set(t,val);
        }else {
            this.setValue(field,cell,t);
        }
    }

    /**
     * @author Jason
     * @date 2020/4/23 14:18
     * @params [method, excelField, cell, val, t]
     * @return void
     * 设值
     */
    private void setValue(Method method,ExcelField excelField,Cell cell,Object val,T t)
            throws InvocationTargetException, IllegalAccessException, ParseException, NoSuchMethodException, InstantiationException {
        if(StringUtil.isNotBlank(excelField.targetMethod())){
            Object target = method.getParameterTypes()[0].newInstance();
            Method targetMethod = target.getClass().getMethod(excelField.targetMethod(), excelField.targetClass());
            if(excelField.useTemplate()){
                targetMethod.invoke(target,val);
            }else{
                this.invoke(targetMethod,cell,target);
            }

            method.invoke(t,target);
        }else {
            if(excelField.useTemplate()){
                method.invoke(t,val);
            }else {
                this.invoke(method,cell,t);
            }
        }
    }

    /**
     * @author Jason
     * @date 2020/4/20 14:32
     * @params [method, cell, instance]
     * 根据不同参数类型执行方法
     * @return void
     */
    private void invoke(Method method,Cell cell,Object instance)
            throws InvocationTargetException, IllegalAccessException, ParseException {

        if(cell == null || cell.toString().length() == 0){
            return;
        }
        //检测excel单元格是否为数字类型
        boolean numberFlag = cell.getCellTypeEnum() == CellType.NUMERIC;
        //检测excel单元格是否为日期类型
        boolean dateFlag = numberFlag && HSSFDateUtil.isCellDateFormatted(cell);

        Class<?> type = method.getParameterTypes()[0];
        //判断对象的方法参数的类型
        if(type == String.class){
            if(dateFlag){
                Date date = cell.getDateCellValue();
                String format = new SimpleDateFormat(ExcelConfig.DATE_IMPORT_FORMAT).format(date);
                method.invoke(instance, format);
            }else if(numberFlag){
                method.invoke(instance,
                        new Double(cell.getNumericCellValue()).intValue()+"");
            }else{
                method.invoke(instance, cell.toString());
            }
        }
        if(type == Date.class){
            Date date = new SimpleDateFormat(ExcelConfig.DATE_IMPORT_FORMAT).parse(cell.toString());
            method.invoke(instance,date);
        }else if(type == Integer.class){
            method.invoke(instance,new Double(cell.toString()).intValue());
        }else if(type == Double.class){
            method.invoke(instance,new Double(cell.toString()));
        }else if(type == Long.class){
            method.invoke(instance,new Double(cell.toString()).longValue());
        }else if(type == Float.class){
            method.invoke(instance,new Double(cell.toString()).floatValue());
        }else if(type == Short.class){
            method.invoke(instance,new Double(cell.toString()).shortValue());
        }else if(type == Byte.class){
            method.invoke(instance,new Double(cell.toString()).byteValue());
        }else if(type == Boolean.class){
            if(ExcelConfig.IMPORT_TRUE.equals(cell.toString())){
                method.invoke(instance,true);
            }else if(ExcelConfig.IMPORT_FALSE.equals(cell.toString())){
                method.invoke(instance,false);
            }
        }
    }

    /**
     * @author Jason
     * @date 2020/4/20 14:32
     * @params [field, cell, instance]
     * 根据不同参数类型给字段设置
     * @return void
     */
    private void setValue(Field field,Cell cell,Object instance) throws IllegalAccessException, ParseException {
        if(cell == null || cell.toString().length() == 0){
            return;
        }
        //检测excel单元格是否为数字类型
        boolean numberFlag = cell.getCellTypeEnum() == CellType.NUMERIC;
        //检测excel单元格是否为日期类型
        boolean dateFlag = numberFlag && HSSFDateUtil.isCellDateFormatted(cell);

        if(field.getType() == String.class){
            if(dateFlag){
                Date date = cell.getDateCellValue();
                String format = new SimpleDateFormat(ExcelConfig.DATE_IMPORT_FORMAT).format(date);
                field.set(instance, format);
            }else if(numberFlag){
                field.set(instance,
                        new Double(cell.getNumericCellValue()).intValue());
            }else{
                field.set(instance, cell.toString());
            }
        }

        if(field.getType() == Date.class){
            Date date = new SimpleDateFormat(ExcelConfig.DATE_IMPORT_FORMAT).parse(cell.toString());
            field.set(instance,date);
        }else if(field.getType() == Integer.class){
            field.set(instance,new Double(cell.getNumericCellValue()).intValue());
        }else if(field.getType() == Double.class){
            field.set(instance,cell.getNumericCellValue());
        }else if(field.getType() == Long.class){
            field.set(instance,new Double(cell.getNumericCellValue()).longValue());
        }else if(field.getType() == Float.class){
            field.set(instance,new Double(cell.getNumericCellValue()).floatValue());
        }else if(field.getType() == Short.class){
            field.set(instance,new Double(cell.getNumericCellValue()).shortValue());
        }else if(field.getType() == Byte.class){
            field.set(instance,new Double(cell.getNumericCellValue()).byteValue());
        }else if(field.getType() == Boolean.class){
            if(ExcelConfig.IMPORT_TRUE.equals(cell.toString())){
                field.set(instance,true);
            }else if(ExcelConfig.IMPORT_FALSE.equals(cell.toString())){
                field.set(instance,false);
            }
        }
    }

    public ExcelImport<T> setStartRow(int startRow) {
        this.startRow = startRow;
        return this;
    }

    public ExcelImport<T> setStartSheet(int startSheet) {
        this.startSheet = startSheet;
        return this;
    }

    public ExcelImport<T> setSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    /**
     * @author Jason
     * @date 2020/3/31 13:50
     * @params [template]
     * 配置模板
     * @return com.jason.util.ExcelImport<T>
     */
    public ExcelImport<T> putTemplate(String key, Map<String, String> template) {
        if(null == this.template){
            this.template = new HashMap<>(10);
        }
        this.template.put(key,template);
        return this;
    }

    /**
     * @author Jason
     * @date 2020/4/1 9:46
     * @params [autoMappingByFieldName]
     * @return com.jason.util.ExcelImport<T>
     * 自动根据字段名映射，默认开启
     */
    public ExcelImport<T> setAutoMappingByFieldName(boolean autoMappingByFieldName) {
        this.autoMappingByFieldName = autoMappingByFieldName;
        return this;
    }

    /**
     * @author Jason
     * @date 2020/4/1 9:47
     * @params []
     * 获取工作簿对象
     * @return org.apache.poi.ss.usermodel.Sheet
     */
    public Sheet getSheet() throws IOException {
        if(initialized || sheet == null){
            this.init();
        }
        return sheet;
    }
}
