package com.jason;

import com.jason.entity.GeneralExcel;
import com.jason.entity.demo.Classes;
import com.jason.entity.demo.Student;
import com.jason.entity.myexport.ExportNoUseAnno;
import com.jason.entity.myexport.ExportUseAnno;
import com.jason.entity.myimport.ImportNoUseAnno;
import com.jason.entity.myimport.ImportUseAnno;
import com.jason.mapper.SimpleMapper;
import com.jason.service.ClassesService;
import com.jason.util.ExcelConfig;
import com.jason.util.ExcelExport;
import com.jason.util.ExcelImport;
import com.jason.util.SqlSessionFactoryUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.util.*;

/**
 * @Auther: Jason
 * @Date: 2019/12/20 23:52
 * @Description:
 */
public class Demo {

    private static int size = 1000;
    private ClassesService classesService = new ClassesService();
    private static List<ExportNoUseAnno> noUseAnnoList = new ArrayList<>(size);
    private static List<ExportUseAnno> useAnnoList = new ArrayList<>(size);
    private static List<GeneralExcel> generalExcelList = new ArrayList<>(size);
    private static String[] headRow;

    //初始化测试数据
    @Before
    public void init(){
        headRow = new String[]{
                "Integer","Long","Short","Byte","Double","Float","aBoolean","Char","parent","字典数据","date"
        };
        for(int i = 0,j=size;i<size;i++,j--){
            ExportNoUseAnno noUseAnno = new ExportNoUseAnno();
            if(i%3 == 0){
                noUseAnno.setaBoolean(true);
            }else{
                noUseAnno.setaBoolean(false);
            }
            noUseAnno.setaByte((byte) i);
            noUseAnno.setaShort((short)j);
            noUseAnno.setaLong((long) 141);
            noUseAnno.setaDouble(1.41);
            noUseAnno.setaFloat((float) 2.23);
            noUseAnno.setaInteger(i);
            noUseAnno.setaCharacter('A');
            noUseAnno.setTemplate(i+"");
            noUseAnno.setDate(new Date());
            noUseAnno.setParent(new ExportNoUseAnno().setTemplate(i+""));
            noUseAnnoList.add(noUseAnno);

            ExportUseAnno useAnno = new ExportUseAnno();
            if(i%3 == 0){
                useAnno.setaBoolean(true);
            }else{
                useAnno.setaBoolean(false);
            }
            useAnno.setaByte((byte) i);
            useAnno.setaShort((short)j);
            useAnno.setaLong((long) 141);
            useAnno.setaDouble(1.41);
            useAnno.setaFloat((float) 2.23);
            useAnno.setaInteger(i);
            useAnno.setaCharacter('A');
            useAnno.setTemplate(i+"");
            useAnno.setDate(new Date());
            ExportUseAnno parent = new ExportUseAnno().setTemplate(i + "");
            Student student = new Student().setName("张" + i);
            student.setClasses(new Classes().setName("141班"));
            parent.setStudent(student);
            useAnno.setParent(parent);
            useAnnoList.add(useAnno);

            GeneralExcel gen = new GeneralExcel();
            if(i%3 == 0){
                gen.setaBoolean(true);
            }else{
                gen.setaBoolean(false);
            }
            gen.setaByte((byte) i);
            gen.setaShort((short)j);
            gen.setaLong((long) 141);
            gen.setaDouble(1.41);
            gen.setaFloat((float) 2.23);
            gen.setaInteger(i);
            gen.setaCharacter('A');
            gen.setDictData1(i+"");
            gen.setDictData2(j+"");
            gen.setDate(new Date());
            GeneralExcel p1 = new GeneralExcel().setDictData1(i + "");
            Student s1 = new Student().setName("张" + i);
            s1.setClasses(new Classes().setName("141班"));
            p1.setStudent(student);
            gen.setStudent(s1);
            gen.setParent(p1);
            generalExcelList.add(gen);
        }
    }

    @Test
    public void test() {
        SimpleMapper simpleMapper = SqlSessionFactoryUtil.getMapper(SimpleMapper.class);
        simpleMapper.DML("delete from classes where id = 123456");
        System.out.println(simpleMapper.getSingleColumnString("select name from classes where id = 1"));
        System.out.println(simpleMapper.selectObject("select * from classes"));
        System.out.println(simpleMapper.selectSingleColumnStringList("select name from classes"));
        System.out.println(simpleMapper.selectSingleObject("select * from classes where id = 1"));
        System.out.println(simpleMapper.selectObject("select * from student"));
    }

    @Test
    public void generalExcel() throws IOException {
        //导出模拟数据
        ExcelExport<GeneralExcel> genExport = new ExcelExport<>(GeneralExcel.class, "通用导入导出");
        genExport.putTemplate("default",ExcelConfig.getTemplateTitle());
        Map<String, String> template1 = new HashMap<>(1);
        template1.put("1","字典数据2");
        genExport.putTemplate("第二个",template1);
        String e1 = genExport.outPutData(generalExcelList);
        genExport.writeToFile("C:/Users/mh262/Desktop/general.xlsx");
        System.out.println(e1);

        //导入
        File file = new File("C:/Users/mh262/Desktop/general.xlsx");
        ExcelImport<GeneralExcel> genImport = new ExcelImport<>(new FileInputStream(file),GeneralExcel.class);
        Map<String, String> template2 = new HashMap<>(1);
        genImport.putTemplate("default",ExcelConfig.getTemplateCode()).putTemplate("第二个",template2);
        template2.put("字典数据2","1");
        List<GeneralExcel> list = genImport.getObjectList();
        String e2 = genImport.getErrorMsg();
        System.out.println(list);
        System.out.println(e2);
    }

    @Test
    //使用注解导出
    //支持模板格式
    public void exportUseAnno() throws InvocationTargetException, IllegalAccessException, IOException {

        ExcelExport<ExportUseAnno> export = new ExcelExport<>(ExportUseAnno.class,"导出测试");
        //配置转换模板格式
        export.putTemplate("default",ExcelConfig.getTemplateTitle());
        Map<String, String> template = new HashMap<>(1);
        template.put("0","第二个map");
        export.putTemplate("第二个",template);
        String errorMsg = export.outPutData(useAnnoList);
        export.writeToFile("C:/Users/mh262/Desktop/exportUseAnno.xlsx");
        System.out.println(errorMsg);

        /*
        ExcelExport<ExcelUseAnno> export = new ExcelExport<>(ExcelUseAnno.class,"导出测试");
        //配置转换模板格式
        export.putTemplate("default",ExcelConfig.getTemplateTitle());
        //第二种写法
        for(int i = 0;i<useAnnoList.size();i++){
            try {
                if(i == 3){
                    throw new RuntimeException("模拟异常回滚");
                }
                export.outPutData(useAnnoList.get(i));
            }catch (Exception e){
                System.out.println("第"+i+"行数据导出失败："+e.getMessage());
            }
        }
        export.writeToFile("C:/Users/user/Desktop/exportUseAnno.xlsx");
        */
    }

    @Test
    //不使用注解导出
    //不支持模板格式
    public void exportNoUseAnno() throws IOException, InvocationTargetException, IllegalAccessException {

        ExcelExport<ExportNoUseAnno> export = new ExcelExport<>(ExportNoUseAnno.class,"导出测试");
        export.putTemplate("default",ExcelConfig.getTemplateTitle());
        //不使用注解时，需传入标题行
        String errorMsg = export.outPutData(noUseAnnoList, headRow);
        export.writeToFile("C:/Users/mh262/Desktop/exportNoUseAnno.xlsx");
        System.out.println(errorMsg);


        /*
        //第二种写法
        ExcelExport<ExcelNoUseAnno> export = new ExcelExport<>(ExcelNoUseAnno.class,"导出测试");
        export.putTemplate("default",ExcelConfig.getTemplateTitle());
        for(int i = 0 ; i < noUseAnnoList.size(); i++){
            try {
                if(i == 3){
                    throw new RuntimeException("模拟异常回滚");
                }
                //不使用注解时，需传入标题行
                export.outPutData(noUseAnnoList.get(i),headRow);
            }catch (Exception e){
                System.out.println("第"+i+"行数据导出失败："+e.getMessage());
            }
        }
        export.writeToFile("C:/Users/user/Desktop/exportNoUseAnno.xlsx");
        */
    }

    @Test
    //使用注解导入
    //支持模板格式
    public void importUseAnno() throws NoSuchMethodException, ParseException, InstantiationException, IOException, IllegalAccessException, InvocationTargetException, NoSuchFieldException {
        File file = new File("C:/Users/mh262/Desktop/exportUseAnno.xlsx");
        ExcelImport<ImportUseAnno> excelImport = new ExcelImport<>(new FileInputStream(file), ImportUseAnno.class);
        //设置模板格式
        excelImport.putTemplate("default",ExcelConfig.getTemplateCode());
        excelImport.putTemplate("parent",ExcelConfig.getTemplateTitle());
        List<ImportUseAnno> list =  excelImport.getObjectList();
        String errorMsg = excelImport.getErrorMsg();
        System.out.println(list);
        System.out.println(errorMsg);

        /*
        //第二种写法
        ExcelImport<ImportUseAnno> excelImport = new ExcelImport<>(new FileInputStream(file), ImportUseAnno.class);
        //设置模板格式
        excelImport.putTemplate(ExcelConfig.getTemplateCode());
        List<ImportUseAnno> list = new ArrayList<>();
        Sheet sheet = excelImport.getSheet();
        //此处i应从1开始
        for(int i = 1; i < sheet.getLastRowNum()+1;i++){
            try {
                if(i == 3){
                    throw new RuntimeException("模拟异常回滚");
                }
                ImportUseAnno useAnno = excelImport.getObject(sheet.getRow(i));
                if(null != useAnno){
                    list.add(useAnno);
                }
            }catch (Exception e){
                System.out.println("第"+i+"行数据导入失败："+e.getMessage());
            }
        }
        System.out.println(list);
        */
    }

    @Test
    //不使用注解导入只能取出与实体类字段名称一样的列
    //不支持模板格式
    public void importNoUseAnno() throws IOException, NoSuchMethodException, InstantiationException, IllegalAccessException, InvocationTargetException, NoSuchFieldException, ParseException {
        File file = new File("C:/Users/user/Desktop/exportUseAnno.xlsx");

        ExcelImport<ImportNoUseAnno> excelImport = new ExcelImport<>(new FileInputStream(file), ImportNoUseAnno.class);
        List<ImportNoUseAnno> list = new ArrayList<>();
        excelImport.getObjects(list);
        System.out.println(list);

        /*
        //第二种写法
        ExcelImport<ImportNoUseAnno> excelImport = new ExcelImport<>(new FileInputStream(file), ImportNoUseAnno.class);
        //设置模板格式
        excelImport.putTemplate(ExcelConfig.getTemplateCode());
        List<ImportNoUseAnno> list = new ArrayList<>();
        Sheet sheet = excelImport.getSheet();
        //此处i应从1开始
        for(int i = 1; i < sheet.getLastRowNum()+1;i++){
            try {
                if(i == 3){
                    throw new RuntimeException("模拟异常回滚");
                }
                ImportNoUseAnno noUseAnno = excelImport.getObject(sheet.getRow(i));
                if(null != noUseAnno){
                    list.add(noUseAnno);
                }
            }catch (Exception e){
                System.out.println("第"+i+"行数据导入失败："+e.getMessage());
            }
        }
        System.out.println(list);
        */
    }
}
