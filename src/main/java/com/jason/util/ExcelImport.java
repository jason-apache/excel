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
        }finally {
            try {
                is.close();
            }catch (IOException e1){
                e1.printStackTrace();
                try {
                    Thread.sleep(500);
                    is.close();
                }catch (Exception e2){
                    e1.printStackTrace();
                }
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
            } catch (IllegalStateException | NumberFormatException | ParseException e){
                String[] msg = e.getMessage().split(":");
                errorMsg.append("错误信息：单元格数据格式异常，第").append(i+this.startRow).append("行，").append(msg[msg.length-1]).append("\r\n");
            } catch (Exception e){
                e.printStackTrace();
                errorMsg.append("错误信息：第").append(i+this.startRow).append("行，").append(e.toString()).append("\r\n");
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
        if(autoMappingByFieldName){
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
     * 设值，过滤模板格式
     */
    private void setValue(Object o,ExcelField excelField, Cell cell, Object t)
            throws IllegalAccessException, ParseException, InvocationTargetException, NoSuchMethodException, InstantiationException, NoSuchFieldException {

        //模板格式转换后的数据传递进去
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

        if(StringUtil.isNotBlank(excelField.call())){
            this.recursion(field,excelField,null,cell,instance,0);
        }else{
            field.setAccessible(true);
            //如果使用了callMethod属性
            if(StringUtil.isNotBlank(excelField.callMethod())){
                Object thisInstance = field.getType().newInstance();
                Method callMethod = field.getType().getMethod(excelField.callMethod(), excelField.callClass());
                this.invoke(callMethod,excelField,cell,thisInstance);
                field.set(instance,thisInstance);
            }else {
                this.setValue(field,excelField,cell,instance);
            }
        }
    }

    /**
    * @author Jason
    * @date 2020/4/26 17:29
    * @params [field, excelField, node, cell, instance, count]
    * 递归调用
    * @return java.lang.Object
    */
    private Object recursion(Field field,ExcelField excelField,String node,Cell cell,Object instance,int count)
            throws ParseException, IllegalAccessException, NoSuchFieldException, InstantiationException,
            NoSuchMethodException, InvocationTargetException {

        field.setAccessible(true);
        if(node == null){
            node = excelField.call();
        }
        //层级调用
        int index = node.indexOf(ExcelConfig.CALL_SEPARATOR);
        //满足递归条件
        if(index != -1){
            //节点拆分
            String curNode = node.substring(0,index);
            String nextNode = node.substring(index+1);
            //当前实例中获取call的当前节点属性
            Field curField = field.getType().getDeclaredField(curNode);
            //创建出下一节点实例
            Object nextInstance = curField.getType().newInstance();
            curField.setAccessible(true);
            //如果是第一次递归，则先把当前属性创建出一个实例，并建立关联，再将这个实例与下一个实例相关联
            if(count == 0){
                Object thisInstance = field.getType().newInstance();
                field.set(instance,thisInstance);
                curField.set(thisInstance,nextInstance);
            }else{
                //否则将当前实例直接关联
                curField.set(instance,nextInstance);
            }
            //记录递归次数
            count++;
            this.recursion(curField,excelField,nextNode,cell,nextInstance,count);
        }else {
            //满足递归退出条件：所有节点实例都创建完毕
            Field curField = field.getType().getDeclaredField(node);
            curField.setAccessible(true);
            //如果使用了callMethod属性
            if(StringUtil.isNotBlank(excelField.callMethod())){
                //如果没有递归，需要让原字段与处理后的对象实例建立关系
                if(count == 0){
                    Object originalInstance = field.getType().newInstance();
                    Object thisInstance = curField.getType().newInstance();
                    Method callMethod = curField.getType().getMethod(excelField.callMethod(), excelField.callClass());
                    this.invoke(callMethod,excelField,cell,thisInstance);
                    curField.set(originalInstance,thisInstance);
                    field.set(instance,originalInstance);
                }else {
                    Object thisInstance = curField.getType().newInstance();
                    Method callMethod = curField.getType().getMethod(excelField.callMethod(), excelField.callClass());
                    this.invoke(callMethod,excelField,cell,thisInstance);
                    curField.set(instance,thisInstance);
                }
            }else{
                //如果没有递归，需要让原字段与处理后的对象实例建立关系
                if(count == 0){
                    Object thisInstance = field.getType().newInstance();
                    this.setValue(curField,excelField,cell,thisInstance);
                    field.set(instance,thisInstance);
                }else{
                    this.setValue(curField,excelField,cell,instance);
                }
            }
        }
        return instance;
    }

    /**
     * @author Jason
     * @date 2020/4/23 14:18
     * @params [method, excelField, cell, val, t]
     * @return void
     * 设值
     */
    private void setMethodValue(Method method,ExcelField excelField,Cell cell,Object t)
            throws InvocationTargetException, IllegalAccessException, ParseException, NoSuchMethodException, InstantiationException, NoSuchFieldException {
        //是否使用call属性
        if(StringUtil.isNotBlank(excelField.call())){
            //获取方法参数列表第一个参数类型
            Class<?> arg = method.getParameterTypes()[0];
            //分隔call节点
            int index = excelField.call().indexOf(ExcelConfig.CALL_SEPARATOR);
            Field field;
            if(index != -1){
                field = arg.getDeclaredField(excelField.call().substring(0, index));
                Object thisInstance = field.getType().newInstance();
                String nextNode = excelField.call().substring(index+1);
                Object instance = this.recursion(field, excelField, nextNode, cell, thisInstance, 0);
                //判断是否引用本身类型
                if(instance.getClass() != arg){
                    Object parameter = arg.newInstance();
                    field.setAccessible(true);
                    field.set(parameter,instance);
                    //递归后返回的实例set至实体中
                    method.invoke(t,parameter);
                }else{
                    method.invoke(t,instance);
                }
            }else {
                if(StringUtil.isNotBlank(excelField.callMethod())){
                    //如果使用了callMethod属性
                    field = arg.getDeclaredField(excelField.call());
                    field.setAccessible(true);
                    Object thisInstance = field.getType().newInstance();
                    Method targetMethod = field.getType().getMethod(excelField.callMethod(), excelField.callClass());
                    this.invoke(targetMethod,excelField,cell,thisInstance);
                    Object parameter = arg.newInstance();
                    field.set(parameter,thisInstance);
                    method.invoke(t,parameter);
                }else{
                    Object parameter = arg.newInstance();
                    field = arg.getDeclaredField(excelField.call());
                    field.setAccessible(true);
                    this.setValue(field,excelField,cell,parameter);
                    method.invoke(t,parameter);
                }
            }
        }else if(StringUtil.isNotBlank(excelField.callMethod())){
            //如果使用了callMethod属性
            Object target = method.getParameterTypes()[0].newInstance();
            Method targetMethod = target.getClass().getMethod(excelField.callMethod(), excelField.callClass());
            this.invoke(targetMethod,excelField,cell,target);
            method.invoke(t,target);
        }else {
            this.invoke(method,excelField,cell,t);
        }
    }

    /**
     * @author Jason
     * @date 2020/4/20 14:32
     * @params [method, cell, instance]
     * 根据不同参数类型执行方法
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
                return;
            }
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
                        new Double(cell.toString()).intValue());
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
     * @params [field, cell, instance]
     * 根据不同参数类型给字段设置
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
                return;
            }
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
                        new Double(cell.toString()).intValue());
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
            field.set(instance,cell.toString());
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
    public ExcelImport<T> putTemplate(String nameKey, Map<String, String> template) {
        if(null == this.template){
            this.template = new HashMap<>(10);
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
}
