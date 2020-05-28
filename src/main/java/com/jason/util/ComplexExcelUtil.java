package com.jason.util;

import com.jason.anno.ExcelField;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.ParseException;
import java.util.*;

/**
 * @author: Jason
 * @Date: 2020/5/27 15:59
 * @Description: 处理复杂性excel需求的工具类
 */
public class ComplexExcelUtil {

    /**
    * @author Jason
    * @date 2020/5/28 11:01
    * @params []
    * 私有化构造方法
    * @return
    */
    private ComplexExcelUtil(){}

    private static final int DEFAULT_SIZE = 16;
    /**
     * class映射关系
     */
    private static final Map<Class<?>,Map<ExcelField,Object>> CLASS_MAPPING = new HashMap<>(DEFAULT_SIZE);
    /**
     * 导出注解排序
     */
    private static final Map<Class<?>, List<ExcelField>> ANNOTATION_MAPPING = new HashMap<>(DEFAULT_SIZE);


    /***********************************导出****************************************/
    /**
    * @author Jason
    * @date 2020/5/28 10:52
    * @params [clazz]
    * 初始化指定类的注解和方法、属性
    * @return void
    */
    private static <T> void initClazz(Class<T> clazz){
        Field[] fields = clazz.getDeclaredFields();
        Method[] methods = clazz.getDeclaredMethods();
        Map<ExcelField, Object> map = null;
        List<ExcelField> list = null;
        for (Field field : fields) {
            ExcelField annotation = field.getAnnotation(ExcelField.class);
            if(null != annotation){
                if(null == map){
                    map = new HashMap<>(fields.length);
                }
                if(null == list){
                    list = new ArrayList<>(fields.length);
                }
                map.put(annotation,field);
                list.add(annotation);
            }
        }
        for (Method method : methods) {
            ExcelField annotation = method.getAnnotation(ExcelField.class);
            if(null != annotation){
                if(null == map){
                    map = new HashMap<>(methods.length);
                }
                if(null == list){
                    list = new ArrayList<>(methods.length);
                }
                map.put(annotation,method);
                list.add(annotation);
            }
        }
        if(null != list){
            list.sort(Comparator.comparingInt(ExcelField::sort));
        }
        CLASS_MAPPING.put(clazz,map);
        ANNOTATION_MAPPING.put(clazz,list);
    }

    /**
    * @author Jason
    * @date 2020/5/28 10:53
    * @params [sheet, config, curRow]
    * 创建标题行，放回当前行数
    * @return int
    */
    public static int createTitle(Sheet sheet,ExportConfig config,int curRow){
        Row row = sheet.createRow(curRow);
        CellStyle cellStyle = config.getStyleMap().get(ExcelConfig.Style.TITLE);
        Cell cell = row.createCell(0);
        ExcelExport.setCellStyle(cell,cellStyle);
        cell.setCellValue(config.getTitle());
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, config.getHeadRow().length-1);
        ExcelExport.setCellStyle(row.createCell(config.getHeadRow().length-1),cellStyle);
        sheet.addMergedRegion(region);
        return ++curRow;
    }

    /**
    * @author Jason
    * @date 2020/5/28 10:54
    * @params [sheet, config, curRow]
    * 创建headRow，放回当前行数
    * @return int
    */
    public static int createHead(Sheet sheet, ExportConfig config,int curRow){
        Row row = sheet.createRow(curRow);
        CellStyle cellStyle = config.getStyleMap().get(ExcelConfig.Style.HEAD_ROW);
        int curCellIndex = 0;
        for (int i = 0; i < config.headRow.length; i++) {
            String t = config.headRow[i];
            Cell cell = row.createCell(curCellIndex);
            cell.setCellValue(t);
            ExcelExport.setCellStyle(cell,cellStyle);
            if(sheet.getColumnWidth(i) < ExcelConfig.Style.CELL_MIN_WIDTH){
                sheet.setColumnWidth(i,ExcelConfig.Style.CELL_MIN_WIDTH);
            }
            curCellIndex++;
        }
        return ++curRow;
    }

    /**
    * @author Jason
    * @date 2020/5/28 10:54
    * @params [clazz, sheet, collection, curRow, template, config]
    * 创建行，放回当前行数
    * @return int
    */
    public static <T> int outputData(Class<T> clazz,Sheet sheet,Collection<T> collection,int curRow,
                                     Map<String,Map<String,String>> template,ExportConfig config)
            throws NoSuchMethodException, NoSuchFieldException, IllegalAccessException, InvocationTargetException {
        if(null != config.getTitle()){
            curRow = createTitle(sheet,config,curRow);
        }
        if(null != config.getHeadRow()){
            curRow = createHead(sheet,config,curRow);
        }
        if(null != collection){
            Map<ExcelField, Object> annotationMapping = CLASS_MAPPING.get(clazz);
            List<ExcelField> annotationList = ANNOTATION_MAPPING.get(clazz);
            //为空则代表第一次加载此类，初始化此类
            if(null == annotationMapping || null == annotationList){
                initClazz(clazz);
                annotationMapping = CLASS_MAPPING.get(clazz);
                annotationList = ANNOTATION_MAPPING.get(clazz);
            }
            int temp = curRow;
            if(null != annotationList && null != annotationMapping){
                for (T t : collection) {
                    Row row = sheet.createRow(curRow);
                    int curCellNum = 0;
                    for (ExcelField excelField : annotationList) {
                        Object o = annotationMapping.get(excelField);
                        if(o instanceof Method && excelField.isImport()){
                            continue;
                        }
                        //调用导出工具类的方法
                        Object val = ExcelExport.getVal(excelField, o, t);
                        Cell cell = row.createCell(curCellNum);
                        //调用导出工具类的方法
                        ExcelExport.setValue(excelField,cell,val, config.getStyleMap(),
                                ExcelConfig.Style.DEFAULT_STYLE,template);
                        curCellNum++;
                        if(temp == curRow){
                            CellAddress address = cell.getAddress();
                            //设置导出列格式要求
                            ExcelExport.createCellFormat(excelField,template,sheet,address.getColumn(),address.getColumn());
                        }
                    }
                    curRow += config.rowInterval;
                }
            }
        }
        return curRow;
    }

    /**
    * @author Jason
    * @date 2020/5/28 10:57
    * @params [sheet, config, body, curRow]
    * 输出string数组
    * @return int
    */
    public static int outputData(Sheet sheet,ExportConfig config,List<List<String>> body,int curRow){
        if(null != config.getTitle()){
            curRow = createTitle(sheet,config,curRow);
        }
        if(null != config.getHeadRow()){
            curRow = createHead(sheet,config,curRow);
        }
        for (List<String> stringList : body) {
            Row row = sheet.createRow(curRow);
            int curCellNum = 0;
            for (String val : stringList) {
                Cell cell = row.createCell(curCellNum++);
                cell.setCellValue(val);
                //调用导出工具类，设置样式
                ExcelExport.setCellStyle(cell,config.getStyleMap().get(ExcelConfig.Style.DEFAULT_STYLE));
            }
            curRow += config.rowInterval;
        }
        return curRow;
    }

    /**
    * @author Jason
    * @date 2020/5/28 10:57
    * @params [clazz]
    * 加载一个class的所有注解，获取title并返回
    * @return java.lang.String[]
    */
    public static <T> String[] getHeadRow(Class<T> clazz){
        List<ExcelField> list = ANNOTATION_MAPPING.get(clazz);
        if(null == list){
            Field[] fields = clazz.getDeclaredFields();
            Method[] methods = clazz.getDeclaredMethods();
            list = new ArrayList<>(fields.length + methods.length);
            for (Field field : fields) {
                ExcelField annotation = field.getAnnotation(ExcelField.class);
                if(null != annotation && !"".equals(annotation.title())){
                    list.add(annotation);
                }
            }
            for (Method method : methods) {
                ExcelField annotation = method.getAnnotation(ExcelField.class);
                if(null != annotation && !"".equals(annotation.title()) && !annotation.isImport()){
                    list.add(annotation);
                }
            }
            list.sort(Comparator.comparingInt(ExcelField::sort));
            ANNOTATION_MAPPING.put(clazz,list);
        }
        String[] headRow = new String[list.size()];
        for (int i = 0; i < headRow.length; i++) {
            headRow[i] = list.get(i).title();
        }
        return headRow;
    }

    /**
    * @author Jason
    * @date 2020/5/28 10:58
    * @params
    * 导出配置类
    * @return
    */
    public static class ExportConfig{
        /**
         * 导出行间隔
         */
        private int rowInterval = 1;
        /**
         * 工作簿，凭借工作簿对象创建style
         */
        private Workbook workbook;
        /**
         * 首列
         */
        private String[] headRow;
        /**
         * 标题行
         */
        private String title;
        /**
         * 样式
         */
        private Map<String,CellStyle> styleMap;

        /**
        * @author Jason
        * @date 2020/5/28 11:00
        * @params [workbook]
        * 构造方法，workbook是必须的
        * @return
        */
        public ExportConfig(Workbook workbook) {
            this.workbook = workbook;
            this.styleMap = ExcelConfig.Style.getStyles(workbook);
        }

        public ExportConfig(Workbook workbook,int rowInterval) {
            this.rowInterval = rowInterval;
            this.workbook = workbook;
            this.styleMap = ExcelConfig.Style.getStyles(workbook);
        }

        public ExportConfig(Workbook workbook, String[] headRow, String title) {
            this.workbook = workbook;
            this.headRow = headRow;
            this.title = title;
            this.styleMap = ExcelConfig.Style.getStyles(workbook);
        }

        public Workbook getWorkbook() {
            return workbook;
        }

        public ExportConfig setWorkbook(Workbook workbook) {
            this.workbook = workbook;
            return this;
        }

        public int getRowInterval() {
            return this.rowInterval;
        }

        public void setRowInterval(int rowInterval) {
            this.rowInterval = rowInterval;
        }

        public String[] getHeadRow() {
            return headRow;
        }

        public ExportConfig setHeadRow(String[] headRow) {
            this.headRow = headRow;
            return this;
        }

        public String getTitle() {
            return title;
        }

        public ExportConfig setTitle(String title) {
            this.title = title;
            return this;
        }

        public Map<String, CellStyle> getStyleMap() {
            return styleMap;
        }

        public ExportConfig setStyleMap(Map<String, CellStyle> styleMap) {
            this.styleMap = styleMap;
            return this;
        }
    }

    /***********************************导出end****************************************/

    /***********************************导入****************************************/

    /**
    * @author Jason
    * @date 2020/5/28 11:00
    * @params [sheet, config]
    * 获取string数组
    * @return java.util.List<java.util.List<java.lang.String>>
    */
    public static List<List<String>> getStringList(Sheet sheet,ImportConfig config){
        return getStringList(sheet,new ArrayList<>(config.getCollectionSize()),config);
    }

    /**
    * @author Jason
    * @date 2020/5/28 11:02
    * @params [sheet, lists, config]
    * 获取string数组
    * @return java.util.List<java.util.List<java.lang.String>>
    */
    public static List<List<String>> getStringList(Sheet sheet,List<List<String>> lists,ImportConfig config){
        if(null != lists){
            for (int i = config.getStartRow(); i < config.getEndRow(); i++) {
                Row row = sheet.getRow(i);
                if(null != row){
                    short lastCellNum = row.getLastCellNum();
                    List<String> stringList = new ArrayList<>(lastCellNum);
                    for (int j = 0; j < lastCellNum; j++) {
                        Cell cell = row.getCell(j);
                        if(null != cell){
                            stringList.add(cell.toString());
                        }
                    }
                    lists.add(stringList);
                }
            }
        }
        return lists;
    }

    /**
    * @author Jason
    * @date 2020/5/28 11:02
    * @params [clazz, sheet, config]
    * 重载方法，获取泛型list集合
    * @return java.util.List<T>
    */
    public static <T> List<T> getObjectList(Class<T> clazz,Sheet sheet,ImportConfig config)
            throws InstantiationException, IllegalAccessException, InvocationTargetException, NoSuchMethodException, ParseException, NoSuchFieldException {
        return (List<T>) getObjects(clazz,new ArrayList<>(config.getCollectionSize()),sheet,config);
    }

    /**
    * @author Jason
    * @date 2020/5/28 11:02
    * @params [clazz, sheet, config]
    * 重载方法，获取泛型set集合
    * @return java.util.Set<T>
    */
    public static <T> Set<T> getObjectSet(Class<T> clazz,Sheet sheet,ImportConfig config)
            throws InstantiationException, IllegalAccessException, InvocationTargetException, NoSuchMethodException, ParseException, NoSuchFieldException {
        return (Set<T>) getObjects(clazz,new HashSet<>(config.getCollectionSize()),sheet,config);
    }

    /**
    * @author Jason
    * @date 2020/5/28 11:03
    * @params [clazz, collection, sheet, config]
    * 根据注解和sheet，反射对象
    * @return java.util.Collection<T>
    */
    public static <T> Collection<T> getObjects(Class<T> clazz,Collection<T> collection,Sheet sheet,ImportConfig config)
            throws IllegalAccessException, InstantiationException, NoSuchMethodException, NoSuchFieldException, ParseException, InvocationTargetException {
        if(null != collection){
            Map<ExcelField, Object> annotationMapping = CLASS_MAPPING.get(clazz);
            //如果为空则代表第一次加载此类，初始化此类
            if(null == annotationMapping){
                initClazz(clazz);
                annotationMapping = CLASS_MAPPING.get(clazz);
            }
            if(null != annotationMapping){
                Row row = sheet.getRow(config.getStartRow());
                short lastCellNum = row.getLastCellNum();
                Map<String, Integer> titleMapping = new HashMap<>(lastCellNum);
                //初始化title映射关系
                for (int i = 0; i < lastCellNum; i++) {
                    Cell cell = row.getCell(i);
                    if(null != cell){
                        titleMapping.put(row.getCell(i).toString(),i);
                    }
                }
                //遍历sheet，反射对象
                for (int i = config.getStartRow()+1; i < config.getEndRow(); i++) {
                    T t = clazz.newInstance();
                    Row dataRow = sheet.getRow(i);
                    for (ExcelField excelField : annotationMapping.keySet()) {
                        Integer index;
                        if(excelField.position() != -1){
                            index = excelField.position();
                        }else{
                            index = titleMapping.get(excelField.title());
                        }
                        if(null != index){
                            Cell cell = dataRow.getCell(index);
                            Object o = annotationMapping.get(excelField);
                            setValue(o,excelField,cell,t,config.getTemplate());
                        }else{
                            continue;
                        }
                    }
                    collection.add(t);
                }
            }
        }
        return collection;
    }

    /**
    * @author Jason
    * @date 2020/5/28 11:05
    * @params [o, excelField, cell, t, template]
    * 设值
    * @return void
    */
    public static void setValue(Object o,ExcelField excelField, Cell cell, Object t,Map<String,Map<String,String>> template)
            throws IllegalAccessException, ParseException, InvocationTargetException, NoSuchMethodException, InstantiationException, NoSuchFieldException {

        if(o instanceof Method){
            if(excelField.isImport()){
                setMethodValue((Method) o,excelField,cell,t,template);
            }
        }else if(o instanceof Field){
            setFiledValue((Field) o,excelField,cell,t,template);
        }
    }

    /**
    * @author Jason
    * @date 2020/5/28 11:05
    * @params [field, excelField, cell, instance, template]
    * 设值
    * @return void
    */
    public static void setFiledValue(Field field,ExcelField excelField,Cell cell,Object instance,
                                     Map<String,Map<String,String>> template)
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
            recursionSet(o,excelField,null,cell,template);
        }else{
            invokeCallMethod(field,excelField,cell,instance,template);
        }
    }

    /**
    * @author Jason
    * @date 2020/5/28 11:05
    * @params [method, excelField, cell, t, template]
    * 设值
    * @return void
    */
    public static void setMethodValue(Method method, ExcelField excelField, Cell cell, Object t,
                                      Map<String,Map<String,String>> template) throws InvocationTargetException,
            IllegalAccessException, ParseException, NoSuchMethodException, InstantiationException, NoSuchFieldException {
        //是否使用call属性
        if(StringUtil.isNotBlank(excelField.call())) {
            //获取方法参数列表第一个参数类型
            Object parameter = method.getParameterTypes()[0].newInstance();
            recursionSet(parameter,excelField,null,cell,template);
            method.invoke(t,parameter);
        }else{
            if(StringUtil.isNotBlank(excelField.callMethod())){
                //获取方法参数列表第一个参数类型
                Object parameter = method.getParameterTypes()[0].newInstance();
                Method callMethod = parameter.getClass().getDeclaredMethod(excelField.callMethod(), excelField.callClass());
                //调用导入工具类的方法
                ExcelImport.invoke(callMethod,excelField,cell,parameter,template);
                method.invoke(t,parameter);
            }else{
                //调用导入工具类的方法
                ExcelImport.invoke(method,excelField,cell,t,template);
            }
        }
    }

    /**
    * @author Jason
    * @date 2020/5/28 11:05
    * @params [instance, excelField, node, cell, template]
    * 递归设值
    * @return java.lang.Object
    */
    public static Object recursionSet(Object instance,ExcelField excelField,String node,Cell cell,
                                      Map<String,Map<String,String>> template) throws NoSuchFieldException,
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
            return recursionSet(curInstance,excelField,nextNode,cell,template);
        }else{
            Field curField = instance.getClass().getDeclaredField(node);
            curField.setAccessible(true);
            invokeCallMethod(curField,excelField,cell,instance,template);
        }
        return instance;
    }

    /**
    * @author Jason
    * @date 2020/5/28 11:05
    * @params [field, excelField, cell, instance, template]
    * 过滤是否使用callMethod属性
    * @return void
    */
    public static void invokeCallMethod(Field field,ExcelField excelField,Cell cell,Object instance,
                                        Map<String,Map<String,String>> template) throws ParseException,
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
            //调用导入工具类的方法
            ExcelImport.invoke(callMethod,excelField,cell,curInstance,template);
        }else{
            //调用导入工具类的方法
            ExcelImport.setValue(field,excelField,cell,instance,template);
        }
    }

    /**
    * @author Jason
    * @date 2020/5/28 11:06
    * @params
    * 导入配置类
    * @return
    */
    public static class ImportConfig{
        /**
         * 起始行
         */
        private int startRow;
        /**
         * 结束行
         */
        private int endRow;
        /**
         * 默认集合大小
         */
        private int collectionSize;
        /**
         * 字典转换
         */
        private Map<String,Map<String,String>> template;

        /**
        * @author Jason
        * @date 2020/5/28 11:07
        * @params [sheet]
        * 构造方法，sheet是必须的，依靠它初始化endRow和collectionSize
        * @return
        */
        public ImportConfig(Sheet sheet){
            this.endRow = sheet.getLastRowNum()+1;
            this.collectionSize = endRow;
        }

        public ImportConfig(Sheet sheet,int startRow){
            this(sheet);
            this.startRow = startRow;
        }

        public ImportConfig(int startRow, int endRow) {
            this.startRow = startRow;
            this.endRow = endRow;
            this.collectionSize = endRow;
        }

        public ImportConfig(int startRow, int endRow, int collectionSize) {
            this.startRow = startRow;
            this.endRow = endRow;
            this.collectionSize = collectionSize;
        }

        public int getStartRow() {
            return startRow;
        }

        public ImportConfig setStartRow(int startRow) {
            this.startRow = startRow;
            return this;
        }

        public int getEndRow() {
            return endRow;
        }

        public ImportConfig setEndRow(int endRow) {
            this.endRow = endRow;
            return this;
        }

        public int getCollectionSize() {
            return collectionSize;
        }

        public ImportConfig setCollectionSize(int collectionSize) {
            this.collectionSize = collectionSize;
            return this;
        }

        public Map<String, Map<String, String>> getTemplate() {
            return template;
        }

        public ImportConfig setTemplate(Map<String, Map<String, String>> template) {
            this.template = template;
            return this;
        }
    }
}