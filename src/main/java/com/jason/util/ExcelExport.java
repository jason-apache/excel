package com.jason.util;

import com.jason.anno.ExcelField;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author Jason
 * @Date: 2019/12/25 15:42
 * @Description:
 */
public class ExcelExport<T> {

    /**
     * 反射对象
     */
    private final Class<T> clazz;
    /**
     * 工作簿对象
     */
    private SXSSFWorkbook sxssfWorkbook;
    /**
     *工作簿对象
     */
    private SXSSFSheet sheet;
    /**
     * 标题行
     */
    private String title;
    /**
     * 首行列明
     */
    private String[] headRow;
    /**
     *是否已创建标题行
     */
    private boolean hasHeadRow;
    /**
     *是否使用excel注解
     */
    private boolean useAnnotation = true;
    /**
     *当前行
     */
    private int curRow = 0;
    /**
     *注解
     */
    private List<ExcelField> annotationList;
    /**
     *注解映射关系
     */
    private Map<ExcelField,Object> annotationMapping;
    /**
     *不使用注解，默认以字段顺序
     */
    private Field[] fields;
    /**
     *模板格式
     */
    private Map<String,Map<String,String>> template;
    /**
     *样式
     */
    private Map<String, CellStyle> styles;
    /**
     *样式key
     */
    private String styleKey = ExcelConfig.Style.DEFAULT_STYLE;

    public ExcelExport(Class<T> clazz,String title){
        this.clazz = clazz;
        this.title = title;
        init();
    }

    /**
     * @author Jason
     * @date 2020/3/26 17:43
     * 初始化方法
     * @params [collection]
     * @return com.jason.util.ExcelExport<T>
     */
    private void init(){
        //获取Class字段
        Field[] fields = this.clazz.getDeclaredFields();
        //获取Class方法
        Method[] methods = this.clazz.getMethods();
        this.fields = fields;
        annotationList = new ArrayList<>(fields.length + methods.length);
        annotationMapping = new HashMap<>(fields.length + methods.length);
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (null != excelField && !excelField.isImport() && StringUtil.isNotBlank(excelField.title())) {
                annotationList.add(excelField);
                annotationMapping.put(excelField, field);
            }
        }
        for (Method method : methods) {
            ExcelField excelField = method.getAnnotation(ExcelField.class);
            if (null != excelField && !excelField.isImport() && StringUtil.isNotBlank(excelField.title())) {
                annotationList.add(excelField);
                annotationMapping.put(excelField, method);
            }
        }
        //排序
        annotationList.sort(Comparator.comparingInt(ExcelField::sort));

        //首行数组
        headRow = new String[annotationList.size()];
        for(int i = 0 ; i < annotationList.size() ; i ++ ){
            headRow[i] = annotationList.get(i).title();
        }

        this.sxssfWorkbook = new SXSSFWorkbook(ExcelConfig.EXPORT_WORK_SIZE);
        this.sheet = sxssfWorkbook.createSheet(ExcelConfig.SHEET_NAME);
    }

    /**
    * @author Jason
    * @date 2020/3/30 16:27
    * @params []
    * 默认样式
    * @return void
    */
    public void defaultStyles(){
        Map<String, CellStyle> styles = new HashMap<>(3);
        CellStyle style = sxssfWorkbook.createCellStyle();

        //设置headRow样式
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font headRowFont = sxssfWorkbook.createFont();
        headRowFont.setFontName(ExcelConfig.Style.FONT_NAME);
        headRowFont.setFontHeightInPoints(ExcelConfig.Style.FONT_HEAD_SIZE);
        headRowFont.setBold(true);
        style.setFont(headRowFont);
        styles.put(ExcelConfig.Style.HEAD_ROW, style);

        //设置标题样式
        style = sxssfWorkbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font titleFont = sxssfWorkbook.createFont();
        titleFont.setFontName(ExcelConfig.Style.FONT_NAME);
        titleFont.setFontHeightInPoints(ExcelConfig.Style.FONT_TITLE_SIZE);
        titleFont.setBold(true);
        style.setFont(titleFont);
        styles.put(ExcelConfig.Style.TITLE, style);

        //设置普通单元格样式
        style = sxssfWorkbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font dataFont = sxssfWorkbook.createFont();
        dataFont.setFontName(ExcelConfig.Style.FONT_NAME);
        dataFont.setFontHeightInPoints(ExcelConfig.Style.FONT_CELL_SIZE);
        style.setFont(dataFont);
        styles.put(ExcelConfig.Style.DEFAULT_STYLE, style);

        this.styles = styles;
    }

    /**
     * @author Jason
     * @date 2020/3/30 13:07
     * @params [row]
     * 创建标题行
     * @return void
     */
    private void createHeadRow(String[] headRow){
        if(headRow == null || headRow.length == 0){
            return;
        }
        if(styles == null){
            defaultStyles();
        }
        //设置标题
        if(StringUtil.isNotBlank(this.title)){
            Row row = this.createRow();
            Cell cell = row.createCell(0);
            cell.setCellValue(this.title);
            cell.setCellStyle(styles.get(ExcelConfig.Style.TITLE));
            CellRangeAddress region = new CellRangeAddress(0, 0, 0, headRow.length-1);
            row.createCell(headRow.length-1).setCellStyle(styles.get(ExcelConfig.Style.DEFAULT_STYLE));
            sheet.addMergedRegion(region);
        }
        Row row = this.createRow();
        int curCellNum = 0;
        for (int i = 0; i < headRow.length; i++) {
            String head = headRow[i];
            if (head != null) {
                Cell cell = row.createCell(curCellNum++);
                cell.setCellValue(head);
                cell.setCellStyle(styles.get(ExcelConfig.Style.HEAD_ROW));
                //下面这行代码容易引起文件受损
                //this.sheet.setColumnWidth((short)1,head.getBytes().length * 2 * 256);
                if(sheet.getColumnWidth(i) < ExcelConfig.Style.CELL_MIN_WIDTH){
                    sheet.setColumnWidth(i,ExcelConfig.Style.CELL_MIN_WIDTH);
                }
            }
        }
        this.headRow = headRow;
        hasHeadRow = true;
    }

    /**
     * @author Jason
     * @date 2020/3/27 10:00
     * @params []
     * 新增一行数据
     * @return void
     */
    private Row createRow(){
        return this.sheet.createRow(curRow++);
    }

    /**
     * @author Jason
     * @date 2020/3/30 13:06
     * @params [row, curCellColumns]
     * 新增一个单元格
     * @return org.apache.poi.ss.usermodel.Cell
     */
    private Cell createCell(Row row,int curCellColumns){
        return row.createCell(curCellColumns);
    }

    /**
    * @author Jason
    * @date 2020/3/26 17:43
    * 输出数据至excel
    * @params [collection]
    * @return com.jason.util.ExcelExport<T>
    */
    public String outPutData(Collection<T> collection){
        StringBuilder errorMsg = new StringBuilder();
        if(null != collection && !collection.isEmpty()){
            int currentElement = 0;
            for(T t : collection){
                try {
                    this.outPutData(t);
                    currentElement++;
                }catch (Exception e){
                    e.printStackTrace();
                    errorMsg.append("错误信息：第").append(currentElement).append("行，").append(e.getMessage()).append("\r\n");
                }
            }
        }
        return errorMsg.toString();
    }

    /**
     * @author Jason
     * @date 2020/3/26 17:43
     * 输出数据至excel
     * @params [collection]
     * @return com.jason.util.ExcelExport<T>
     */
    public ExcelExport<T> outPutData(T t) throws IllegalAccessException,
            InvocationTargetException, NoSuchFieldException, NoSuchMethodException {

        //创建首行
        if(!hasHeadRow || curRow == 0){
            this.createHeadRow(this.headRow);
        }

        int curCellNum = 0;

        Row row = this.createRow();

        //是否使用注解
        if(useAnnotation){
            //开始创建数据
            for (ExcelField excelField : annotationList) {
                Object o = annotationMapping.get(excelField);
                Cell cell = this.createCell(row, curCellNum++);
                Object val = this.getVal(excelField, o, t);
                this.setValue(excelField,cell, val);
            }
        }else{
            for(int i = 0; i < headRow.length; i++){
                if(i >= fields.length){
                    break;
                }
                Field field = fields[i];
                field.setAccessible(true);
                Object val = field.get(t);
                Cell cell = this.createCell(row, curCellNum++);
                this.setValue(null,cell,val);
            }
        }

        return this;
    }

    /**
    * @author Jason
    * @date 2020/4/23 11:12
    * @params [excelField, o, t]
    * @return java.lang.Object
    * 获取值
    */
    private Object getVal(ExcelField excelField,Object o,T t) throws InvocationTargetException,
            IllegalAccessException, NoSuchFieldException, NoSuchMethodException {
        if(null == excelField || null == o){
            return null;
        }
        if(o instanceof Method){
            return this.getMethodVal(excelField,(Method)o,t);
        }else if(o instanceof Field){
            return this.getFieldVal(excelField,(Field) o,t);
        }else {
            return null;
        }
    }

    /**
    * @author Jason
    * @date 2020/4/23 11:06
    * @params [excelField, method, t]
    * @return java.lang.Object
    * 获取值
    */
    private Object getMethodVal(ExcelField excelField,Method method,T t) throws InvocationTargetException,
            IllegalAccessException {
        return method.invoke(t);
    }

    /**
    * @author Jason
    * @date 2020/4/23 11:07
    * @params [excelField, field, t]
    * @return java.lang.Object
    * 获取值
    */
    private Object getFieldVal(ExcelField excelField,Field field,T t) throws IllegalAccessException,
            NoSuchFieldException, NoSuchMethodException, InvocationTargetException {
        if(StringUtil.isNotBlank(excelField.call())){
            //把当前属性get出来直接丢入递归
            field.setAccessible(true);
            Object o = field.get(t);
            return this.recursionGet(o,excelField,null);
        }else {
            field.setAccessible(true);
            return field.get(t);
        }
    }

    /**
    * @author Jason
    * @date 2020/4/28 9:34
    * @params [instance, excelField, node]
    * 递归获取值
    * @return java.lang.Object
    */
    private Object recursionGet(Object instance,ExcelField excelField,String node)
            throws NoSuchFieldException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        if(node == null){
            node = excelField.call();
        }
        if(instance == null){
            return null;
        }
        int index = node.indexOf(ExcelConfig.CALL_SEPARATOR);
        if(index != -1){
            //节点拆分
            String curNode = node.substring(0,index);
            String nextNode = node.substring(index+1);
            //获取当前节点的字段
            Field curField = instance.getClass().getDeclaredField(curNode);
            curField.setAccessible(true);
            //取出，丢入递归
            Object curInstance = curField.get(instance);
            return this.recursionGet(curInstance,excelField,nextNode);
        }else{
            Field curField = instance.getClass().getDeclaredField(node);
            curField.setAccessible(true);
            //判断是否使用callMethod属性
            if(StringUtil.isNotBlank(excelField.callMethod())){
                //先把当前属性从实例中取出
                Object curInstance = curField.get(instance);
                //调用callMethod
                Method callMethod = curInstance.getClass().getDeclaredMethod(excelField.callMethod());
                return callMethod.invoke(curInstance);
            }else{
                //取出该属性直接返回
                return curField.get(instance);
            }
        }
    }

    /**
    * @author Jason
    * @date 2020/3/31 11:50
    * @params [collection, headRow]
    * 不使用注解 返回错误信息
    * @return String
    */
    public String outPutData(Collection<T> collection, String[] headRow){
        StringBuilder errorMsg = new StringBuilder();
        if(null != collection && !collection.isEmpty()){
            int currentElement = 0;
            for(T t : collection){
                try {
                    this.outPutData(t,headRow);
                    currentElement++;
                }catch (Exception e){
                    errorMsg.append("错误信息：第").append(currentElement).append("行，").append(e.getMessage()).append("\r\n");
                }
            }
        }
        return errorMsg.toString();
    }

    /**
    * @author Jason
    * @date 2020/3/31 11:47
    * @params [t, headRow]
    * 不使用注解
    * @return com.jason.util.ExcelExport<T>
    */
    public ExcelExport<T> outPutData(T t, String[] headRow) throws InvocationTargetException,
            IllegalAccessException, NoSuchFieldException, NoSuchMethodException {
        this.useAnnotation = false;
        if(!hasHeadRow || curRow == 0){
            this.createHeadRow(headRow);
        }
        return this.outPutData(t);
    }

    /**
     * @author Jason
     * @date 2020/3/26 17:43
     * 输出excel至流
     * @params [collection]
     * @return com.jason.util.ExcelExport<T>
     */
    public ExcelExport<T> write(OutputStream os) throws IOException {
        sxssfWorkbook.write(os);
        return this;
    }

    /**
     * @author Jason
     * @date 2020/3/26 17:43
     * 输出excel至客户端
     * @params [collection]
     * @return com.jason.util.ExcelExport<T>
     */
    public ExcelExport<T> write(HttpServletResponse response, String fileName) throws IOException {
        response.reset();
        response.setContentType("application/octet-stream; charset=utf-8");
        response.setHeader("Content-Disposition", "attachment; filename="+fileName);
        this.write(response.getOutputStream());
        return this;
    }

    /**
     * @author Jason
     * @date 2020/3/26 17:43
     * 输出excel至文件
     * @params [collection]
     * @return com.jason.util.ExcelExport<T>
     */
    public ExcelExport<T> writeToFile(String fileName) throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(fileName);
        this.write(fileOutputStream);
        return this;
    }

    /**
     * @author Jason
     * @date 2020/3/26 17:43
     * 输出excel至文件
     * @params [collection]
     * @return com.jason.util.ExcelExport<T>
     */
    public ExcelExport<T> writeToFile(File file) throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        this.write(fileOutputStream);
        return this;
    }

    /**
     * @author Jason
     * @date 2020/3/30 13:17
     * @params [excelField,cell, field]
     * 设值至单元格，转换模板数据
     * @return void
     */
    private void setValue(ExcelField excelField,Cell cell,Object object){
        //设置样式
        cell.setCellStyle(styles.get(styleKey));
        if(object == null){
            cell.setCellValue("");
            return;
        }
        if(null != excelField && excelField.useTemplate()){
            Map<String, String> map = template.get(excelField.templateNameKey());
            if(null != map){
                cell.setCellValue(map.get(object.toString()));
                return;
            }
        }

        if(object instanceof String){
            cell.setCellValue(object.toString());
        }else if(object instanceof Integer){
            cell.setCellValue((Integer) object);
        }else if(object instanceof Long){
            cell.setCellValue((Long) object);
        }else if(object instanceof Double){
            cell.setCellValue((Double) object);
        }else if(object instanceof Character){
            cell.setCellValue(object.toString());
        }else if(object instanceof Short){
            cell.setCellValue((Short) object);
        }else if(object instanceof Byte){
            cell.setCellValue((Byte) object);
        }else if(object instanceof Float){
            //解决float精度问题
            cell.setCellValue(Double.parseDouble(object.toString()));
        }else if(object instanceof Boolean){
            if((Boolean) object){
                cell.setCellValue(ExcelConfig.EXPORT_TRUE);
            }else {
                cell.setCellValue(ExcelConfig.EXPORT_FALSE);
            }
        }else if(object instanceof Date){
            SimpleDateFormat format = new SimpleDateFormat(ExcelConfig.DATE_EXPORT_FORMAT);
            cell.setCellValue(format.format(object));
        }else{
            cell.setCellValue(object.toString());
        }
    }

    /**
    * @author Jason
    * @date 2020/3/30 15:28
    * @params [template]
    * 设值模板格式
    * @return com.jason.util.ExcelExport<T>
    */
    public ExcelExport<T> putTemplate(String nameKey,Map<String, String> template) {
        if(this.template == null){
            this.template = new HashMap<>(16);
        }
        this.template.put(nameKey,template);
        return this;
    }

    public ExcelExport<T> setHeadRow(String[] headRow) {
        this.headRow = headRow;
        return this;
    }

    /**
    * @author Jason
    * @date 2020/3/30 15:44
    * @params [styles]
    * 设置样式
    * @return com.jason.util.ExcelExport<T>
    */
    public ExcelExport<T> setStyles(Map<String, CellStyle> styles) {
        this.styles = styles;
        return this;
    }

    /**
    * @author Jason
    * @date 2020/3/30 16:14
    * @params [styleKey]
    * 设置样式key
    * @return com.jason.util.ExcelExport<T>
    */
    public ExcelExport<T> setStyleKey(String styleKey) {
        this.styleKey = styleKey;
        return this;
    }
}
