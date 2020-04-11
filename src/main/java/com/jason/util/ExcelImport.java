package com.jason.util;

import com.jason.anno.ExcelField;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
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

    private int startRow;
    private int startSheet;
    private String sheetName;
    private Class<T> clazz;
    private InputStream is;
    private Method[] methods;
    private XSSFWorkbook xssfWorkbook;
    private HSSFWorkbook hssfWorkbook;
    private byte[] bytes;
    private Sheet sheet;
    private boolean initialized;
    private boolean useTemplate;
    private boolean autoMappingByFieldName = true;

    //方法名与excel列的映射关系，title对应excel的列名
    private Map<String,String> methodMapping = new HashMap<>();

    //参数位置与方法名的映射关系
    private Map<String,Integer> parameterMapping = new HashMap<>();

    //参数类型与方法名的映射关系
    private Map<String,Class<?>> argsTypeMapping = new HashMap<>();

    //当使用引用类型时，当前类方法与目标类方法名的映射
    private Map<String,String> targetMethodMapping = new HashMap<>();

    //当使用引用类型时，目标类方法参数类型的映射
    private Map<String,Class<?>> targetFieldTypeMapping = new HashMap<>();

    //不声明方法式设值时，默认以字段名映射excel
    private Set<String> fieldsSet;

    //不声明方法式设值时，字段与excel的映射关系
    private Map<String,Integer> fieldPositionMapping = new HashMap<>();

    //使用模板格式列
    private Set<String> templateColumn = new HashSet<>();
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
        this.methods = clazz.getDeclaredMethods();
        this.is = is;
        this.initMethods();
    }

    private void initMethods(){
        Field[] fields = clazz.getDeclaredFields();
        if(autoMappingByFieldName){
            fieldsSet = new HashSet<>();
            for (int i = 0;i < fields.length; i++){
                Field field = fields[i];
                fieldsSet.add(field.getName());
            }
        }
        for(int i=0;i<methods.length;i++){
            Method method = methods[i];
            ExcelField excelField = method.getAnnotation(ExcelField.class);
            if(excelField == null || !excelField.isImport()){
                continue;
            }
            if(0 <= excelField.position()){
                //设置位置映射关系
                parameterMapping.put(method.getName(),excelField.position());
                argsTypeMapping.put(method.getName(),method.getParameterTypes()[0]);
                if(excelField.useTemplate()){
                    templateColumn.add(method.getName());
                }
            }else if(!"".equals(excelField.title())){
                //设置方法名与列名映射
                methodMapping.put(excelField.title(),method.getName());
                argsTypeMapping.put(method.getName(),method.getParameterTypes()[0]);
                if(excelField.useTemplate()){
                    templateColumn.add(method.getName());
                }
            }
            //引用类型目标方法映射
            if(!"".equals(excelField.targetMethod())){
                targetMethodMapping.put(method.getName(),excelField.targetMethod());
                targetFieldTypeMapping.put(method.getName(),excelField.targetClass());
            }
        }
    }

    private void init() throws IOException {

        try {
            bytes = new byte[is.available()];
            is.read(bytes);
            xssfWorkbook = new XSSFWorkbook(new ByteArrayInputStream(bytes));
            //初始化
            if(StringUtil.isNotBlank(sheetName)){
                sheet = xssfWorkbook.getSheet(sheetName);
            }else{
                sheet = xssfWorkbook.getSheetAt(startSheet);
            }
        }catch (Exception e){
            System.out.println("格式不匹配");
            hssfWorkbook = new HSSFWorkbook(new ByteArrayInputStream(bytes));
            //初始化
            if(StringUtil.isNotBlank(sheetName)){
                sheet = hssfWorkbook.getSheet(sheetName);
            }else{
                sheet = hssfWorkbook.getSheetAt(startSheet);
            }
        }

        //取excel首行列名
        Row firstRow = sheet.getRow(startRow);
        //根据方法名映射，取出excel对应列的位置index，放入参数映射
        for(int i=0;i<firstRow.getLastCellNum();i++){
            String data = firstRow.getCell(i) + "";
            if(methodMapping.containsKey(data)){
                parameterMapping.put(methodMapping.get(data),i);
                //优先以注解映射，注解无法映射再通过字段名映射
                if(useTemplate && (templateColumn.contains(data) || templateColumn.contains(methodMapping.get(data)))){
                    templateColumn.add(i+"");
                }
            }else if(autoMappingByFieldName && fieldsSet.contains(data)){
                fieldPositionMapping.put(data,i);
            }
        }
        this.initialized = true;
    }

    /**
     * @author Jason
     * @date 2020/3/31 13:21
     * @params [file, startRow, startSheet, collection]
     * 转为Java对象
     * @return java.util.Collection<T>
     */
    public Collection<T> getObjects(Collection<T> collection)
            throws InvocationTargetException, NoSuchMethodException, NoSuchFieldException, InstantiationException,
            IllegalAccessException, IOException, ParseException {

        if(!this.initialized){
            this.init();
        }

        for(int i=this.startRow+1;i<this.sheet.getLastRowNum()+1;i++){
            T o = this.getObject(this.sheet.getRow(i));
            if(null != o){
                collection.add(o);
            }
        }

        return collection;
    }

    /**
    * @author Jason
    * @date 2020/3/30 17:18
    * @params [row]
    * 解析excel
    * @return T
    */
    public T getObject(Row row) throws IllegalAccessException,
            NoSuchFieldException, InstantiationException, NoSuchMethodException, InvocationTargetException, IOException, ParseException {
            if(null == row){
                return null;
            }
            if(!this.initialized){
                this.init();
            }
            T o = (T) clazz.newInstance();
            //根据参数位置映射，开始解析excel
            for (String s : parameterMapping.keySet()) {
                //根据方法名、参数类型取出对象对应方法
                Method method = clazz.getDeclaredMethod(s, argsTypeMapping.get(s));
                //取出单元格
                Integer index = parameterMapping.get(s);
                Cell cell = row.getCell(index);

                //存在引用类型时，先试图设值引用对象的属性，再set到实体中
                String typeMethod = targetMethodMapping.get(s);
                if(StringUtil.isNotBlank(typeMethod)){
                    Object target = argsTypeMapping.get(s).newInstance();
                    Method targetMethod = target.getClass().getMethod(typeMethod, targetFieldTypeMapping.get(s));
                    Class<?> type = targetFieldTypeMapping.get(s);
                    //从模板中获取
                    if(useTemplate && templateColumn.contains(index+"") && type == String.class){
                        String val = template.get(cell + "");
                        targetMethod.invoke(target,val);
                    }else{
                        this.invoke(targetMethod,cell,type,target);
                    }
                    //set到实体中
                    method.invoke(o,target);
                }else{
                    //取出对应参数类型
                    Class<?> type = argsTypeMapping.get(s);

                    //从模板中获取
                    if(useTemplate && templateColumn.contains(index+"") && type == String.class){
                        String val = template.get(cell + "");
                        method.invoke(o,val);
                    }else{
                        this.invoke(method,cell,type,o);
                    }
                }
            }
            //如果对象中字段值为空，则代表注解映射失败，尝试按照字段名设值
            for (String s : fieldPositionMapping.keySet()){
                Field field = o.getClass().getDeclaredField(s);
                //获取私有属性访问权
                field.setAccessible(true);
                if(null == field.get(o)){
                    Cell cell = row.getCell(fieldPositionMapping.get(field.getName()));
                    Class<?> type = field.getType();
                    this.setValue(field,cell,type,o);
                }
            }

            return o;
    }


    //根据参数类型设值
    private void invoke(Method method,Cell cell,Class<?> type,Object instance)
            throws InvocationTargetException, IllegalAccessException, ParseException {

        if(cell == null || type == null || (cell+"").length() == 0){
            return;
        }
        //检测excel单元格是否为数字类型
        boolean numberFlag = cell.getCellTypeEnum() == CellType.NUMERIC;
        //检测excel单元格是否为日期类型
        boolean dateFlag = numberFlag && HSSFDateUtil.isCellDateFormatted(cell);

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
                method.invoke(instance, cell+"");
            }
        }
        if(type == Date.class){
            Date date = new SimpleDateFormat(ExcelConfig.DATE_IMPORT_FORMAT).parse(cell+"");
            method.invoke(instance,date);
        }else if(numberFlag && type == Integer.class){
            method.invoke(instance,new Double(cell.getNumericCellValue()).intValue());
        }else if(numberFlag && type == Double.class){
            method.invoke(instance,cell.getNumericCellValue());
        }else if(numberFlag && type == Long.class){
            method.invoke(instance,new Double(cell.getNumericCellValue()).longValue());
        }else if(numberFlag && type == Float.class){
            method.invoke(instance,new Double(cell.getNumericCellValue()).floatValue());
        }else if(numberFlag && type == Short.class){
            method.invoke(instance,new Double(cell.getNumericCellValue()).shortValue());
        }else if(numberFlag && type == Byte.class){
            method.invoke(instance,new Double(cell.getNumericCellValue()).byteValue());
        }else if(type == Boolean.class){
            if(ExcelConfig.IMPORT_TRUE.equals(cell+"")){
                method.invoke(instance,true);
            }else if(ExcelConfig.IMPORT_FALSE.equals(cell+"")){
                method.invoke(instance,false);
            }
        }
    }

    private void setValue(Field field,Cell cell,Class<?> type,Object instance) throws IllegalAccessException, ParseException {
        if(cell == null || type == null || (cell+"").length() == 0){
            return;
        }
        //检测excel单元格是否为数字类型
        boolean numberFlag = cell.getCellTypeEnum() == CellType.NUMERIC;
        //检测excel单元格是否为日期类型
        boolean dateFlag = numberFlag && HSSFDateUtil.isCellDateFormatted(cell);

        if(type == String.class){
            if(dateFlag){
                Date date = cell.getDateCellValue();
                String format = new SimpleDateFormat(ExcelConfig.DATE_IMPORT_FORMAT).format(date);
                field.set(instance, format);
            }else if(numberFlag){
                field.set(instance,
                        new Double(cell.getNumericCellValue()).intValue()+"");
            }else{
                field.set(instance, cell+"");
            }
        }

        if(type == Date.class){
            Date date = new SimpleDateFormat(ExcelConfig.DATE_IMPORT_FORMAT).parse(cell+"");
            field.set(instance,date);
        }else if(numberFlag && type == Integer.class){
            field.set(instance,new Double(cell.getNumericCellValue()).intValue());
        }else if(numberFlag && type == Double.class){
            field.set(instance,cell.getNumericCellValue());
        }else if(numberFlag && type == Long.class){
            field.set(instance,new Double(cell.getNumericCellValue()).longValue());
        }else if(numberFlag && type == Float.class){
            field.set(instance,new Double(cell.getNumericCellValue()).floatValue());
        }else if(numberFlag && type == Short.class){
            field.set(instance,new Double(cell.getNumericCellValue()).shortValue());
        }else if(numberFlag && type == Byte.class){
            field.set(instance,new Double(cell.getNumericCellValue()).byteValue());
        }else if(type == Boolean.class){
            if(ExcelConfig.IMPORT_TRUE.equals(cell+"")){
                field.set(instance,true);
            }else if(ExcelConfig.IMPORT_FALSE.equals(cell+"")){
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
    public ExcelImport<T> addTemplate(Map<String, String> template) {
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
