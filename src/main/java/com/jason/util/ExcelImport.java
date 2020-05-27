package com.jason.util;

import com.jason.anno.ExcelField;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

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
    private Integer startRow;
    /**
     * 解析工作簿
     */
    private Integer sheetIndex;
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
     * 工作簿
     */
    private Sheet sheet;
    private Workbook workbook;
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
    /**
     * 错误信息
     */
    private String errorMsg;

    /**
    * @author Jason
    * @date 2020/5/19 11:03
    * @params [is, clazz]
    * 构造方法，默认使用XSSFWorkbook对象
    * @return
    */
    public ExcelImport(InputStream is,Class<T> clazz){
        ExcelField field = clazz.getAnnotation(ExcelField.class);
        //根据注解中的属性设初值
        if(null != field){
            this.startRow = field.startRow() > 0 ? field.startRow() - 1 : 0;
            this.sheetIndex = field.sheetIndex() > 0 ? field.sheetIndex() - 1 : 0;
            this.sheetName = field.sheetName();
        }
        this.clazz = clazz;
        this.is = is;
        this.initMethod();
    }

    /**
     * @author Jason
     * @date 2020/5/25 9:53
     * @params [is, clazz, sheetIndex, startRow]
     * 重载构造方法
     * @return
     */
    public ExcelImport(InputStream is,Class<T> clazz,Integer sheetIndex,Integer startRow){
        this(is,clazz,null,sheetIndex,startRow);
    }

    /**
     * @author Jason
     * @date 2020/5/25 9:53
     * @params [is, clazz, sheetIndex, startRow]
     * 重载构造方法
     * @return
     */
    public ExcelImport(InputStream is,Class<T> clazz,String sheetName,Integer startRow){
        this(is,clazz,sheetName,null,startRow);
    }

    /**
     * @author Jason
     * @date 2020/5/25 9:53
     * @params [is, clazz, sheetIndex, startRow]
     * 重载构造方法
     * @return
     */
    public ExcelImport(InputStream is,Class<T> clazz,String sheetName,Integer sheetIndex,Integer startRow){
        this.is = is;
        this.clazz = clazz;
        this.sheetName = sheetName;
        this.sheetIndex = sheetIndex == null ? 0 : sheetIndex;
        this.startRow = startRow;
        this.initMethod();
    }

    /**
     * @author Jason
     * @date 2020/4/22 10:18
     * @params []
     * 初始化工作薄
     * @return void
     */
    private void init(){
        try {
            this.workbook = WorkbookFactory.create(this.is);
            if(StringUtil.isNotBlank(sheetName)){
                sheet = workbook.getSheet(sheetName);
            }else{
                sheet = workbook.getSheetAt(sheetIndex);
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
        }catch (IOException | InvalidFormatException e){
            e.printStackTrace();
        }
    }

    /**
     * @author Jason
     * @date 2020/4/22 10:19
     * @params []
     * 初始化方法
     * @return void
     */
    private void initMethod(){

        Field[] fields = clazz.getDeclaredFields();
        Method[] methods = clazz.getDeclaredMethods();
        annotationList = new ArrayList<>(fields.length + methods.length);
        annotationMapping = new HashMap<>(fields.length + methods.length);
        //自动根据字段名映射
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (null != excelField && StringUtil.isNotBlank(excelField.title())) {
                annotationList.add(excelField);
                annotationMapping.put(excelField, field);
            } else {
                if (autoMappingByFieldName) {
                    if (null == fieldsSet) {
                        fieldsSet = new HashSet<>(fields.length);
                    }
                    fieldsSet.add(field);
                }
            }
        }
        for (Method method : methods) {
            ExcelField excelField = method.getAnnotation(ExcelField.class);
            if (null != excelField && excelField.isImport() && StringUtil.isNotBlank(excelField.title())) {
                annotationList.add(excelField);
                annotationMapping.put(excelField, method);
            }
        }
    }

    /**
    * @author Jason
    * @date 2020/5/25 9:22
    * @params []
    * 转为Java对象 返回错误信息
    * @return java.lang.String
    */
    public List<T> getObjectList() throws IOException {
        return this.getObjectList(ExcelConfig.DEFAULT_COLLECTION_SIZE);
    }

    /**
    * @author Jason
    * @date 2020/5/25 9:28
    * @params [size]
    * 重载
    * @return java.lang.String
    */
    public List<T> getObjectList(int size) throws IOException {
        return (List<T>) this.getObjects(new ArrayList<>(size));
    }

    /**
    * @author Jason
    * @date 2020/5/25 9:24
    * @params []
    * 转为Java对象 返回错误信息
    * @return java.lang.String
    */
    public Set<T> getObjectSet() throws IOException {
        return this.getObjectSet(ExcelConfig.DEFAULT_COLLECTION_SIZE);
    }

    /**
    * @author Jason
    * @date 2020/5/25 9:28
    * @params [size]
    * 重载
    * @return java.lang.String
    */
    public Set<T> getObjectSet(int size) throws IOException {
        return (Set<T>) this.getObjects(new HashSet<>(size));
    }

    /**
     * @author Jason
     * @date 2020/3/31 13:21
     * @params [file, startRow, startSheet, collection]
     * 转为Java对象 返回错误信息
     * @return String
     */
    public Collection<T> getObjects(Collection<T> collection) throws IOException {

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
            } catch (IllegalStateException | NumberFormatException | ParseException e){
                String[] msg = e.getMessage().split(":");
                errorMsg.append("错误信息：单元格数据格式异常，第").append(i+this.startRow).append("行，").append(msg[msg.length-1]).append("\r\n");
            } catch (Exception e){
                e.printStackTrace();
                errorMsg.append("错误信息：第").append(i+this.startRow).append("行，").append(e.toString()).append("\r\n");
            }
            this.errorMsg = errorMsg.toString();
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
    public T getObject(Row row) throws IllegalAccessException, InstantiationException,
            NoSuchMethodException, InvocationTargetException, IOException, ParseException, NoSuchFieldException {
        if(null == row || row.getLastCellNum() == 0){
            return null;
        }
        if(!this.initialized){
            this.init();
        }
        T t = clazz.newInstance();
        //根据参数位置映射，开始解析excel
        for (ExcelField excelField : annotationList) {
            Object o = annotationMapping.get(excelField);
            Cell cell = null;
            //如果使用了position属性
            if (excelField.position() != -1) {
                cell = row.getCell(excelField.position());
            } else {
                Integer index = titleMapping.get(excelField.title());
                if (null != index) {
                    cell = row.getCell(index);
                } else {
                    continue;
                }
            }

            this.setValue(o,excelField,cell,t);
         }
        //是否自动根据参数名映射 默认开启
        if(autoMappingByFieldName && null != fieldsSet){
            for(Field f : fieldsSet){
                f.setAccessible(true);
                Integer index = titleMapping.get(f.getName());
                if(null != index){
                    if(null == f.get(t)){
                        Cell cell = row.getCell(index);
                        this.setValue(f,null,cell,t);
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
     * 设值
     */
    private void setValue(Object o,ExcelField excelField, Cell cell, Object t)
            throws IllegalAccessException, ParseException, InvocationTargetException, NoSuchMethodException, InstantiationException, NoSuchFieldException {

        if(o instanceof Method){
            this.setMethodValue((Method) o,excelField,cell,t);
        }else if(o instanceof Field){
            this.setFiledValue((Field) o,excelField,cell,t);
        }
    }

    /**
     * @author Jason
     * @date 2020/4/23 14:17
     * @params [field, excelField, cell, val, t]
     * @return void
     * 设值
     */
    private void setFiledValue(Field field,ExcelField excelField,Cell cell,Object instance)
            throws ParseException, IllegalAccessException, NoSuchFieldException, InstantiationException,
            NoSuchMethodException, InvocationTargetException {

        field.setAccessible(true);
        if(StringUtil.isNotBlank(excelField.call())){
            Object o;
            //如果该属性为空，则new一个并建立关联
            if(field.get(instance) == null){
                o = field.getType().newInstance();
                field.set(instance,o);
            }else{
                o = field.get(instance);
            }
            this.recursionSet(o,excelField,null,cell);
        }else{
            this.invokeCallMethod(field,excelField,cell,instance);
        }
    }

    /**
     * @author Jason
     * @date 2020/4/23 14:18
     * @params [method, excelField, cell, val, t]
     * @return void
     * 设值
     */
    private void setMethodValue(Method method,ExcelField excelField,Cell cell,Object t) throws InvocationTargetException,
            IllegalAccessException, ParseException, NoSuchMethodException, InstantiationException, NoSuchFieldException {
        //是否使用call属性
        if(StringUtil.isNotBlank(excelField.call())) {
            //获取方法参数列表第一个参数类型
            Object parameter = method.getParameterTypes()[0].newInstance();
            this.recursionSet(parameter,excelField,null,cell);
            method.invoke(t,parameter);
        }else{
            if(StringUtil.isNotBlank(excelField.callMethod())){
                //获取方法参数列表第一个参数类型
                Object parameter = method.getParameterTypes()[0].newInstance();
                Method callMethod = parameter.getClass().getDeclaredMethod(excelField.callMethod(), excelField.callClass());
                this.invoke(callMethod,excelField,cell,parameter);
                method.invoke(t,parameter);
            }else{
                this.invoke(method,excelField,cell,t);
            }
        }
    }

    /**
    * @author Jason
    * @date 2020/4/29 10:07
    * @params [instance, excelField, node, cell]
    * 递归设值
    * @return java.lang.Object
    */
    private Object recursionSet(Object instance,ExcelField excelField,String node,Cell cell) throws NoSuchFieldException,
            IllegalAccessException, ParseException, InstantiationException, NoSuchMethodException, InvocationTargetException {
        if(node == null){
            node = excelField.call();
        }
        if(instance == null){
            return null;
        }
        //层级调用
        int index = node.indexOf(ExcelConfig.CALL_SEPARATOR);
        //满足递归条件
        if(index != -1){
            //节点拆分
            String curNode = node.substring(0,index);
            String nextNode = node.substring(index+1);
            //当前实例中获取call的当前节点属性
            Field curField = instance.getClass().getDeclaredField(curNode);
            curField.setAccessible(true);
            Object curInstance;
            //如果该属性为空，则new一个并建立关联
            if(curField.get(instance) == null){
                curInstance = curField.getType().newInstance();
                curField.set(instance,curInstance);
            }else{
                curInstance = curField.get(instance);
            }
            return this.recursionSet(curInstance,excelField,nextNode,cell);
        }else{
            Field curField = instance.getClass().getDeclaredField(node);
            curField.setAccessible(true);
            this.invokeCallMethod(curField,excelField,cell,instance);
        }
        return instance;
    }

    /**
    * @author Jason
    * @date 2020/4/29 10:04
    * @params [field, excelField, cell, instance]
    * 过滤是否使用callMethod属性
    * @return void
    */
    private void invokeCallMethod(Field field,ExcelField excelField,Cell cell,Object instance) throws ParseException,
            IllegalAccessException, InstantiationException, NoSuchMethodException, InvocationTargetException {
        field.setAccessible(true);
        //如果使用callMethod属性,取出并调用
        if(StringUtil.isNotBlank(excelField.callMethod())){
            Object curInstance;
            if(field.get(instance) == null){
                curInstance = field.getType().newInstance();
                field.set(instance,curInstance);
            }else{
                curInstance = field.get(instance);
            }
            //调用
            Method callMethod = field.getType().getDeclaredMethod(excelField.callMethod(), excelField.callClass());
            this.invoke(callMethod,excelField,cell,curInstance);
        }else{
            this.setValue(field,excelField,cell,instance);
        }
    }

    /**
     * @author Jason
     * @date 2020/4/20 14:32
     * @params [method, excelField, cell, instance]
     * 根据不同参数类型执行方法，转换模板数据
     * @return void
     */
    private void invoke(Method method,ExcelField excelField,Cell cell,Object instance)
            throws InvocationTargetException, IllegalAccessException, ParseException {

        if(cell == null || cell.toString().length() == 0){
            return;
        }
        if(null != excelField && excelField.useTemplate()){
            Map<String, String> map = template.get(excelField.templateNameKey());
            if(null != map){
                String val = map.get(cell.toString());
                method.invoke(instance,val);
            }
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
                        String.valueOf(new Double(cell.toString()).intValue()));
            }else{
                method.invoke(instance, cell.toString());
            }
        }
        if(type == Date.class){
            Date date;
            if(dateFlag){
                date = cell.getDateCellValue();
            }else {
                date = new SimpleDateFormat(ExcelConfig.DATE_IMPORT_FORMAT).parse(cell.toString());
            }
            method.invoke(instance,date);
        }else if(type == Integer.class){
            method.invoke(instance,new Double(cell.toString()).intValue());
        }else if(type == Double.class){
            method.invoke(instance, Double.parseDouble(cell.toString()));
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
        }else if(type == Character.class){
            method.invoke(instance,cell.toString().toCharArray()[0]);
        }
    }

    /**
     * @author Jason
     * @date 2020/4/20 14:32
     * @params [field, excelField, cell, instance]
     * 根据不同参数类型给字段设置，转换模板数据
     * @return void
     */
    private void setValue(Field field,ExcelField excelField,Cell cell,Object instance) throws IllegalAccessException, ParseException {
        if(cell == null || cell.toString().length() == 0){
            return;
        }
        if(null != excelField && excelField.useTemplate()){
            Map<String, String> map = template.get(excelField.templateNameKey());
            if(null != map){
                String val = map.get(cell.toString());
                field.set(instance,val);
            }
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
                        String.valueOf(new Double(cell.toString()).intValue()));
            }else{
                field.set(instance, cell.toString());
            }
        }

        if(field.getType() == Date.class){
            Date date;
            if(dateFlag){
                date = cell.getDateCellValue();
            }else {
                date = new SimpleDateFormat(ExcelConfig.DATE_IMPORT_FORMAT).parse(cell.toString());
            }
            field.set(instance,date);
        }else if(field.getType() == Integer.class){
            field.set(instance,new Double(cell.toString()).intValue());
        }else if(field.getType() == Double.class){
            field.set(instance,Double.parseDouble(cell.toString()));
        }else if(field.getType() == Long.class){
            field.set(instance,new Double(cell.toString()).longValue());
        }else if(field.getType() == Float.class){
            field.set(instance,new Double(cell.toString()).floatValue());
        }else if(field.getType() == Short.class){
            field.set(instance,new Double(cell.toString()).shortValue());
        }else if(field.getType() == Byte.class){
            field.set(instance,new Double(cell.toString()).byteValue());
        }else if(field.getType() == Boolean.class){
            if(ExcelConfig.IMPORT_TRUE.equals(cell.toString())){
                field.set(instance,true);
            }else if(ExcelConfig.IMPORT_FALSE.equals(cell.toString())){
                field.set(instance,false);
            }
        }else if(field.getType() == Character.class){
            field.set(instance,cell.toString().toCharArray()[0]);
        }
    }

    /**
    * @author Jason
    * @date 2020/5/25 10:52
    * @params [template]
    * 重载，配置字典数据转换
    * @return com.jason.util.ExcelImport<T>
    */
    public ExcelImport<T> putTemplate(Map<String, String> template) {
        return this.putTemplate(ExcelConfig.DEFAULT_NAME_KEY,template);
    }

    /**
     * @author Jason
     * @date 2020/3/31 13:50
     * @params [template]
     * 配置字典数据转换
     * @return com.jason.util.ExcelImport<T>
     */
    public ExcelImport<T> putTemplate(String nameKey, Map<String, String> template) {
        if(null == this.template){
            this.template = new HashMap<>();
        }
        this.template.put(nameKey,template);
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

    /**
    * @author Jason
    * @date 2020/5/27 13:12
    * @params []
    * 获取工作簿对象
    * @return org.apache.poi.ss.usermodel.Workbook
    */
    public Workbook getWorkbook() {
        return workbook;
    }

    /**
    * @author Jason
    * @date 2020/5/25 13:19
    * @params []
    * 导入时的错误信息
    * @return java.lang.String
    */
    public String getErrorMsg() {
        return errorMsg;
    }
}
