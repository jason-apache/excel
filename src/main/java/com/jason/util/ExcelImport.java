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

    //解析起始行
    private int startRow;
    //解析工作簿
    private int startSheet;
    //工作簿名称
    private String sheetName;
    //实体类型
    private final Class<T> clazz;
    //输入流
    private final InputStream is;
    //字节数组，保存流
    private byte[] bytes;

    private Sheet sheet;
    //是否初始化
    private boolean initialized;
    //是否使用模板转换
    private boolean useTemplate;
    //自动根据字段名称映射
    private boolean autoMappingByFieldName = true;

    //注解
    private final List<ExcelField> annotationList = new ArrayList<>();
    //注解映射关系
    private final Map<ExcelField,Object> annotationMapping = new HashMap<>();
    //title映射关系
    private final Map<String,Integer> titleMapping = new HashMap<>();

    //不声明方法式设值时，默认以字段名映射excel
    private Set<Field> fieldsSet;

    //模板格式
    private Map<String,String> template;

    public ExcelImport(InputStream is,Class<T> clazz){
        ExcelField field = clazz.getAnnotation(ExcelField.class);
        //根据注解中的属性设初值
        if(null != field){
            startRow = field.startRow() > 0 ? field.startRow() - 1 : 0;
            startSheet = field.startSheet() > 0 ? field.startSheet() - 1 : 0;
            sheetName = "".equals(field.sheetName().trim()) ? null : field.sheetName();
        }
        this.clazz = clazz;
        this.is = is;
        this.initMethods();
    }

    private void initMethods(){
        Field[] fields = clazz.getDeclaredFields();
        Method[] methods = clazz.getDeclaredMethods();
        //自动根据字段名映射
        for(int i = 0; i< methods.length; i ++ ){
            Method method = methods[i];
            ExcelField excelField = method.getAnnotation(ExcelField.class);
            if(null != excelField && excelField.isImport() && StringUtil.isNotBlank(excelField.title())){
                annotationList.add(excelField);
                annotationMapping.put(excelField,method);
            }
        }
        for(int i = 0 ; i < fields.length ; i ++ ){
            Field field = fields[i];
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if(null != excelField && excelField.isImport() && StringUtil.isNotBlank(excelField.title())){
                annotationList.add(excelField);
                annotationMapping.put(excelField,field);
            }else {
                if(autoMappingByFieldName){
                    if(null == fieldsSet){
                        fieldsSet = new HashSet<>();
                    }
                    fieldsSet.add(field);
                }
            }
        }
    }

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
        //取出excel列的位置index，放入title映射
        for(int i=0;i<firstRow.getLastCellNum();i++){
            String data = firstRow.getCell(i) + "";
            titleMapping.put(data,i);
        }
        this.initialized = true;
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
        if(null == row){
            return null;
        }
        if(!this.initialized){
            this.init();
        }
        T t = (T) clazz.newInstance();
        int size = annotationList.size();
        //根据参数位置映射，开始解析excel
        for(int i = 0 ; i < size ; i ++ ){
            ExcelField excelField = annotationList.get(i);
            if(null != excelField){
                Object o = annotationMapping.get(excelField);
                Cell cell = null;
                //如果使用了position属性
                if(excelField.position() != -1){
                    cell = row.getCell(excelField.position());
                }else {
                    cell = row.getCell(titleMapping.get(excelField.title()));
                }
                //方法上的注解
                if(o instanceof Method){
                    //使用了目标方法
                    if(StringUtil.isNotBlank(excelField.targetMethod())){
                        Object target = ((Method) o).getParameterTypes()[0].newInstance();
                        Method targetMethod = target.getClass().getMethod(excelField.targetMethod(), excelField.targetClass());
                        //使用模板
                        if(useTemplate && excelField.useTemplate()){
                            String val = template.get(cell + "");
                            targetMethod.invoke(target,val);
                        }else{
                            this.invoke(targetMethod,cell,target);
                        }
                        //set到实体中
                        ((Method) o).invoke(t,target);
                    }else {
                        //使用模板
                        if(useTemplate && excelField.useTemplate()){
                            String val = template.get(cell + "");
                            ((Method) o).invoke(t,val);
                        }else{
                            this.invoke(((Method) o),cell,t);
                        }
                    }
                    //字段上的注解
                }else if(o instanceof Field){
                    ((Field) o).setAccessible(true);
                    if(useTemplate && excelField.useTemplate()){
                        String val = template.get(cell + "");
                        ((Field) o).set(t,val);
                    }else{
                        this.setValue(((Field) o),cell,t);
                    }
                }
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
            Date date = new SimpleDateFormat(ExcelConfig.DATE_IMPORT_FORMAT).parse(cell+"");
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
            Date date = new SimpleDateFormat(ExcelConfig.DATE_IMPORT_FORMAT).parse(cell+"");
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
    public ExcelImport<T> setTemplate(Map<String, String> template) {
        this.template = template;
        this.useTemplate = true;
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
