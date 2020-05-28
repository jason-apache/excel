package com.jason.util;

import com.jason.anno.ExcelField;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

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
    private Workbook workbook;
    /**
     *工作簿对象
     */
    private Sheet sheet;
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
     *注解，按照list中注解的顺序导出
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
    /**
     * 使用xls后缀，此种方式的局限就是导出的行数至多为65535行，超出65536条后系统就会报错。
     */
    private boolean useHSSF = false;
    /**
     * 使用xlsx后缀，没有65535行限制
     */
    private boolean useSXSSF = true;

    /**
    * @author Jason
    * @date 2020/5/19 13:08
    * @params [clazz, title]
    * 构造方法，默认使用SXSSFworkbook对象
    * @return
    */
    public ExcelExport(Class<T> clazz,String title){
        this(clazz,title,true);
    }

    /**
    * @author Jason
    * @date 2020/5/19 13:09
    * @params [clazz, title, useSXSSF]
    * 构造方法，设置工作簿类型
    * @return
    */
    public ExcelExport(Class<T> clazz,String title,boolean useSXSSF){
        this.useSXSSF = useSXSSF;
        this.useHSSF = !useSXSSF;
        if(useSXSSF){
            this.workbook = new SXSSFWorkbook(ExcelConfig.EXPORT_WORK_SIZE);
        }else{
            this.workbook = new HSSFWorkbook();
        }
        this.sheet = workbook.createSheet(ExcelConfig.SHEET_NAME);
        this.clazz = clazz;
        this.title = title;
        this.init();
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
            if (null != excelField && StringUtil.isNotBlank(excelField.title())) {
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
    }

    /**
    * @author Jason
    * @date 2020/3/30 16:27
    * @params []
    * 默认样式
    * @return void
    */
    public void defaultStyles(){
        this.styles = ExcelConfig.Style.getStyles(this.workbook);
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
            setCellStyle(cell,styles.get(ExcelConfig.Style.TITLE));
            CellRangeAddress region = new CellRangeAddress(0, 0, 0, headRow.length-1);
            setCellStyle(row.createCell(headRow.length-1),styles.get(ExcelConfig.Style.DEFAULT_STYLE));
            sheet.addMergedRegion(region);
        }
        Row row = this.createRow();
        int curCellNum = 0;
        for (int i = 0; i < headRow.length; i++) {
            String head = headRow[i];
            if (head != null) {
                Cell cell = row.createCell(curCellNum++);
                cell.setCellValue(head);
                setCellStyle(cell,styles.get(ExcelConfig.Style.HEAD_ROW));
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
     * @date 2020/5/27 15:38
     * @params [excelField, cell]
     * 设置导出列约束
     * @return void
     */
    private void createCellFormat(ExcelField excelField,Cell cell){
        if(excelField.useTemplate() && this.template != null){
            CellAddress address = cell.getAddress();
            createCellFormat(excelField,this.template,this.sheet,address.getColumn(),address.getColumn());
        }
    }

    /**
     * @author Jason
     * @date 2020/5/27 15:39
     * @params [excelField, template, sheet, startCol, endCol]
     * 静态公有方法
     * @return void
     */
    public static void createCellFormat(ExcelField excelField,Map<String,Map<String,String>> template,Sheet sheet,int startCol,int endCol) {
        if(excelField.useTemplate() && template != null){
            Map<String, String> map = template.get(excelField.templateNameKey());
            if(null != map && !map.isEmpty()){
                Object[] values =  map.values().toArray();
                String[] val = new String[values.length];
                for (int i = 0; i < values.length; i++) {
                    val[i] = values[i].toString();
                }
                DataValidationHelper dataValidationHelper = sheet.getDataValidationHelper();
                CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(0,ExcelConfig.EXPORT_FORMAT_RANGE,startCol,endCol);
                DataValidationConstraint constraint = dataValidationHelper.createExplicitListConstraint(val);
                DataValidation dataValidation = dataValidationHelper.createValidation(constraint, cellRangeAddressList);
                //处理Excel兼容性问题
                if(dataValidation instanceof XSSFDataValidation) {
                    dataValidation.setSuppressDropDownArrow(true);
                    dataValidation.setShowErrorBox(true);
                }else {
                    dataValidation.setSuppressDropDownArrow(false);
                }
                sheet.addValidationData(dataValidation);
            }
        }
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
                    currentElement++;
                    this.outPutData(t);
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
                Object val = getVal(excelField, o, t);
                this.setValue(excelField,cell, val);
                if(this.curRow == 1){
                    this.createCellFormat(excelField,cell);
                }
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
    * 获取值，静态公有方法
    */
    public static Object getVal(ExcelField excelField,Object o,Object instance) throws InvocationTargetException,
            IllegalAccessException, NoSuchFieldException, NoSuchMethodException {
        if(null == excelField || null == o){
            return null;
        }
        if(o instanceof Method){
            return getMethodVal(excelField,(Method)o,instance);
        }else if(o instanceof Field){
            return getFieldVal(excelField,(Field) o,instance);
        }else {
            return null;
        }
    }

    /**
    * @author Jason
    * @date 2020/4/23 11:06
    * @params [excelField, method, t]
    * @return java.lang.Object
    * 获取值，静态公有方法
    */
    public static Object getMethodVal(ExcelField excelField,Method method,Object instance) throws InvocationTargetException,
            IllegalAccessException {
        return method.invoke(instance);
    }

    /**
    * @author Jason
    * @date 2020/4/23 11:07
    * @params [excelField, field, t]
    * @return java.lang.Object
    * 获取值，静态公有方法
    */
    public static Object getFieldVal(ExcelField excelField,Field field,Object instance) throws IllegalAccessException,
            NoSuchFieldException, NoSuchMethodException, InvocationTargetException {
        if(StringUtil.isNotBlank(excelField.call())){
            //把当前属性get出来直接丢入递归
            field.setAccessible(true);
            Object o = field.get(instance);
            return recursionGet(o,excelField,null);
        }else {
            field.setAccessible(true);
            return field.get(instance);
        }
    }

    /**
    * @author Jason
    * @date 2020/4/28 9:34
    * @params [instance, excelField, node]
    * 递归获取值，静态公有方法
    * @return java.lang.Object
    */
    public static Object recursionGet(Object instance,ExcelField excelField,String node)
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
            return recursionGet(curInstance,excelField,nextNode);
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
     * @date 2020/5/27 15:46
     * @params [excelField, cell, object]
     * 设值至单元格，转换模板数据
     * @return void
     */
    private void setValue(ExcelField excelField,Cell cell,Object object){
        setValue(excelField, cell, object,this.styles,this.styleKey,this.template);
    }

    /**
     * @author Jason
     * @date 2020/3/30 13:17
     * @params [excelField,cell, field]
     * 静态公有方法
     * @return void
     */
    public static void setValue(ExcelField excelField,Cell cell,Object object,
                                Map<String,CellStyle> styles,String styleKey,Map<String,Map<String,String>> template){
        //设置样式
        setCellStyle(cell,styles.get(styleKey));
        if(object == null){
            cell.setCellValue("");
            return;
        }
        if(null != excelField && excelField.useTemplate() && null != template){
            Map<String, String> map = template.get(excelField.templateNameKey());
            if(null != map){
                cell.setCellValue(map.get(object.toString()));
            }
            return;
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
            //不属于上面中的类型，则调用toString方法
            cell.setCellValue(object.toString());
        }
    }

    /**
     * @author Jason
     * @date 2020/5/19 13:08
     * @params [cell, cellStyle]
     * 静态公有方法
     * @return void
     */
    public static void setCellStyle(Cell cell,CellStyle cellStyle){
        if(cell instanceof SXSSFCell || cell instanceof XSSFCell){
            cell.setCellStyle(cellStyle);
        }else{
            cell.getCellStyle().cloneStyleFrom(cellStyle);
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
                    currentElement++;
                    this.outPutData(t,headRow);
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
        workbook.write(os);
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
    * @date 2020/5/25 10:51
    * @params [template]
    * 重载，配置字典数据转换
    * @return com.jason.util.ExcelExport<T>
    */
    public ExcelExport<T> putTemplate(Map<String, String> template) {
        return this.putTemplate(ExcelConfig.DEFAULT_NAME_KEY,template);
    }

    /**
    * @author Jason
    * @date 2020/3/30 15:28
    * @params [template]
    * 配置字典数据转换
    * @return com.jason.util.ExcelExport<T>
    */
    public ExcelExport<T> putTemplate(String nameKey,Map<String, String> template) {
        if(this.template == null){
            this.template = new HashMap<>();
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

    /**
    * @author Jason
    * @date 2020/5/27 13:14
    * @params []
    * 获取工作簿对象
    * @return org.apache.poi.ss.usermodel.Workbook
    */
    public Workbook getWorkbook() {
        return workbook;
    }

    /**
    * @author Jason
    * @date 2020/5/27 13:15
    * @params []
    * 获取工作簿对象
    * @return org.apache.poi.ss.usermodel.Sheet
    */
    public Sheet getSheet() {
        return sheet;
    }
}
