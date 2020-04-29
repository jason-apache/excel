# Excel
# java operation Excel 

工作中遇到excel导入导出业务要求，研究了一下poi，故此写了excel工具类

使用poi版本3.17

针对导入导出一些细节性的说明：

import、export不支持单元格合并

import和export均支持字段和方法的注解，但存在一些差异，在方法上加注解时，import调用set方法，export调用get方法

import的注解中（call、callMethod）属性支持方法和字段

export的注解中（call、callMethod）属性仅支持字段，export的方法上面加call本身就是无意义的

在字段上面加注解时，可以同时满足导入与导出，在方法上加注解时，需要指定isImport = true 或者 isImport = false

import调用顺序：先执行字段上的注解，再执行方法