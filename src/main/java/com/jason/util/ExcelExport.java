package com.jason.util;

import com.jason.anno.ExcelField;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Auther: Jason
 * @Date: 2019/12/25 15:42
 * @Description:
 */
public class ExcelExport<T> {

    private Class<T> clazz;

    //工作簿对象
    private SXSSFWorkbook sxssfWorkbook;
    //工作簿对象
    private SXSSFSheet sheet;

    //首行标题
    private String[] headRow;
    //是否已创建标题行
    private boolean hasHeadRow;
    //是否使用excel注解
    private boolean useAnnotation = true;
    //当前行
    private int curRow = 0;

    //字段映射
    private Map<String,Field> filedMapping = new HashMap<>();
    //不使用注解，默认以字段顺序
    private Field[] fields;
    //方法映射
    private Map<String,Method> methodMapping = new HashMap<>();
    //排序映射
    private Map<String,Integer> sortMapping = new HashMap<>();
    //是否使用模板
    private boolean useTemplate;
    //使用模板格式列
    private Map<String,Integer> templateColumn = new HashMap<>();
    //模板格式
    private List<Map<String,String>> template;
    //样式
    private Map<String, CellStyle> styles;
    //样式key
    private String styleKey = ExcelConfig.Style.DEFAULT_STYLE;

    public ExcelExport(Class<T> clazz){
        this.clazz = clazz;
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
        int max = 0;
        //获取Class字段
        Field[] fields = this.clazz.getDeclaredFields();
        this.fields = fields;
        //排序
        Arrays.sort(fields, Comparator.comparingInt(o -> o.getAnnotation(ExcelField.class) == null ? -1 : o.getAnnotation(ExcelField.class).sort()));
        for(int i = 0; i < fields.length; i++){
            ExcelField excelField = fields[i].getAnnotation(ExcelField.class);
            if(null != excelField && StringUtil.isNotBlank(excelField.excelTile()) && !excelField.isImport()){
                filedMapping.put(excelField.excelTile(),fields[i]);
                sortMapping.put(excelField.excelTile(),excelField.sort());
                if(excelField.sort() > max){
                    max = excelField.sort();
                }
                if(excelField.useTemplate()){
                    templateColumn.put(excelField.excelTile(),excelField.templatePosition());
                }
            }
        }
        //获取Class方法
        Method[] methods = this.clazz.getMethods();
        //排序
        Arrays.sort(methods, Comparator.comparingInt(o -> o.getAnnotation(ExcelField.class) == null ? -1 : o.getAnnotation(ExcelField.class).sort()));
        for(int i = 0 ; i < methods.length; i++){
            ExcelField excelField = methods[i].getAnnotation(ExcelField.class);
            if(null != excelField && StringUtil.isNotBlank(excelField.excelTile()) && !excelField.isImport()){
                methodMapping.put(excelField.excelTile(),methods[i]);
                sortMapping.put(excelField.excelTile(),excelField.sort());
                if(excelField.sort() > max){
                    max = excelField.sort();
                }
                if(excelField.useTemplate()){
                    templateColumn.put(excelField.excelTile(),excelField.templatePosition());
                }
            }
        }

        if(max <= 0){
            max = sortMapping.size();
        }
        //首行数组
        headRow = new String[max];
        //根据排序设置首行顺序
        for(String s : sortMapping.keySet()){
            Integer sort = sortMapping.get(s);
            headRow[sort <= 0 ? 0 : sort-1] = s;
        }

        this.sxssfWorkbook = new SXSSFWorkbook();
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
        Map<String, CellStyle> styles = new HashMap<>();
        CellStyle style = sxssfWorkbook.createCellStyle();

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        Font titleFont = sxssfWorkbook.createFont();
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        titleFont.setFontName(ExcelConfig.Style.FONT_NAME);
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBold(true);
        style.setFont(titleFont);
        styles.put(ExcelConfig.Style.HEAD_ROW, style);

        style = sxssfWorkbook.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font dataFont = sxssfWorkbook.createFont();
        dataFont.setFontName(ExcelConfig.Style.FONT_NAME);
        dataFont.setFontHeightInPoints((short) 12);
        style.setFont(dataFont);
        styles.put(ExcelConfig.Style.DEFAULT_STYLE, style);

        this.styles = styles;
    }

    /**
    * @author Jason
    * @date 2020/3/26 17:43
    * 输出数据至excel
    * @params [collection]
    * @return com.jason.util.ExcelExport<T>
    */
    public ExcelExport<T> outPutData(Collection<T> collection) throws IllegalAccessException, InvocationTargetException {
        if(null != collection && !collection.isEmpty()){
            for(T t : collection){
                this.outPutData(t);
            }
        }
        return this;
    }

    /**
     * @author Jason
     * @date 2020/3/26 17:43
     * 输出数据至excel
     * @params [collection]
     * @return com.jason.util.ExcelExport<T>
     */
    public ExcelExport<T> outPutData(T t) throws IllegalAccessException, InvocationTargetException {

        //创建首行
        if(!hasHeadRow || curRow == 0){
            this.createHeadRow(this.headRow);
        }

        int hLength = headRow.length;
        int curCellNum = 0;

        Row row = this.createRow();

        //是否使用注解
        if(useAnnotation){
            //开始创建数据
            for(int i = 0 ; i < hLength; i++){
                String head = headRow[i];
                //优先字段上的注解
                if(null != head){
                    Field field = filedMapping.get(head);
                    if(null != field){
                        field.setAccessible(true);
                        //如果使用了模板格式
                        if(useTemplate && templateColumn.containsKey(head)){
                            //获取模板位置
                            Integer index = templateColumn.get(head);
                            if(index > template.size()-1){
                                index = template.size()-1;
                            }
                            String val = template.get(index).get(field.get(t) + "");
                            Cell cell = this.createCell(row, curCellNum++);
                            cell.setCellStyle(styles.get(styleKey));
                            cell.setCellValue(val);
                        }else{
                            Cell cell = this.createCell(row, curCellNum++);
                            cell.setCellStyle(styles.get(styleKey));
                            this.setValue(cell,field,t);
                        }
                    }else{
                        //其次是方法上的注解
                        Method method = methodMapping.get(head);
                        if(null != method){
                            //如果使用了模板格式
                            if(useTemplate && templateColumn.containsKey(head)){
                                //获取模板位置
                                Integer index = templateColumn.get(head);
                                if(index > template.size()-1){
                                    index = template.size()-1;
                                }
                                String val = template.get(index).get(method.invoke(t) + "");
                                Cell cell = this.createCell(row, curCellNum++);
                                cell.setCellStyle(styles.get(styleKey));
                                cell.setCellValue(val);
                            }else{
                                Object object = method.invoke(t);
                                Cell cell = this.createCell(row, curCellNum++);
                                cell.setCellStyle(styles.get(styleKey));
                                this.setValue(cell,object);
                            }
                        }
                    }
                }
            }
        }else{
            for(int i = 0; i < headRow.length; i++){
                if(i >= fields.length){
                    break;
                }
                Field field = fields[i];
                if(null != field){
                    field.setAccessible(true);
                    Cell cell = this.createCell(row, curCellNum++);
                    cell.setCellStyle(styles.get(styleKey));
                    this.setValue(cell,field,t);
                }
            }
        }

        return this;
    }

    /**
    * @author Jason
    * @date 2020/3/31 11:50
    * @params [collection, headRow]
    * 不使用注解
    * @return com.jason.util.ExcelExport<T>
    */
    public ExcelExport<T> outPutData(Collection<T> collection,String[] headRow)
            throws InvocationTargetException, IllegalAccessException {
        if(null != collection && !collection.isEmpty()){
            for(T t : collection){
                this.outPutData(t,headRow);
            }
        }
        return this;
    }

    /**
    * @author Jason
    * @date 2020/3/31 11:47
    * @params [t, headRow]
    * 不使用注解
    * @return com.jason.util.ExcelExport<T>
    */
    public ExcelExport<T> outPutData(T t,String[] headRow) throws InvocationTargetException, IllegalAccessException {
        this.useAnnotation = false;
        if(!hasHeadRow || curRow == 0){
            this.createHeadRow(headRow);
        }
        return this.outPutData(t);
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
        Row row = this.createRow();
        int curCellNum = 0;
        String[] newHead = null;
        if(useAnnotation){
            newHead = new String[this.sortMapping.size()];
        }
        //循环体外获得长度以提升效率
        int length = headRow.length;
        for(int i = 0 ; i < length ; i++){
            String head = headRow[i];
            if(head != null){
                if(null != newHead){
                    newHead[curCellNum] = head;
                }
                Cell cell = row.createCell(curCellNum++);
                cell.setCellStyle(styles.get(ExcelConfig.Style.HEAD_ROW));
                cell.setCellValue(head);
                this.sheet.setColumnWidth((short)i,head.getBytes().length * 2 * 256);
            }
        }
        if(null != newHead){
            this.headRow = newHead;
        }else{
            this.headRow = headRow;
        }
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
     * @params [cell, field, t]
     * 设值至单元格
     * @return void
     */
    private void setValue(Cell cell,Field field,T t) throws IllegalAccessException {

        Object obj = field.get(t);
        this.setValue(cell,obj);
    }

    /**
     * @author Jason
     * @date 2020/3/30 13:17
     * @params [cell, field, t]
     * 设值至单元格
     * @return void
     */
    private void setValue(Cell cell,Object object){
        if(object == null){
            cell.setCellValue("");
            return;
        }

        if(object instanceof String){
            cell.setCellValue(object+"");
        }else if(object instanceof Integer){
            cell.setCellValue((Integer) object);
        }else if(object instanceof Long){
            cell.setCellValue((Long) object);
        }else if(object instanceof Double){
            cell.setCellValue((Double) object);
        }else if(object instanceof Character){
            cell.setCellValue(object+"");
        }else if(object instanceof Short){
            cell.setCellValue((Short) object);
        }else if(object instanceof Byte){
            cell.setCellValue((Byte) object);
        }else if(object instanceof Float){
            cell.setCellValue((Float) object);
        }else if(object instanceof Boolean){
            if((Boolean) object){
                cell.setCellValue(ExcelConfig.EXPORT_TRUE);
            }else {
                cell.setCellValue(ExcelConfig.EXPORT_FALSE);
            }
        }else if(object instanceof Date){
            SimpleDateFormat format = new SimpleDateFormat(ExcelConfig.DATE_EXPORT_FORMAT);
            cell.setCellValue(format.format(object));
        }
    }

    /**
    * @author Jason
    * @date 2020/3/30 15:28
    * @params [template]
    * 设值模板格式
    * @return com.jason.util.ExcelExport<T>
    */
    public ExcelExport<T> addTemplate(Map<String, String> template) {
        if(this.template == null){
            this.template = new ArrayList<>();
        }
        this.template.add(template);
        this.useTemplate = true;
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
