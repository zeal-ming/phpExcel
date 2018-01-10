<?php

error_reporting(E_ALL);
 
date_default_timezone_set('Asia/ShangHai');
 
/** PHPExcel_IOFactory */
require_once './PHPExcel.php';
 
 $filename = "student/student.xls";
// Check prerequisites
if (!file_exists($filename)) {
	exit("not found student.xls.\n");
}
 
 //自动获取文件的类型
$fileType = PHPExcel_IOFactory::identify($filename);
$reader = PHPExcel_IOFactory::createReader($fileType); //设置以Excel5格式(Excel97-2003工作簿)
$PHPExcel = $reader->load($filename); // 载入excel文件
$sheet = $PHPExcel->getSheet(0); // 读取第一個工作表
$highestRow = $sheet->getHighestRow(); // 取得总行数
$highestColumm = $sheet->getHighestColumn(); // 取得总列数
$highestColumm= PHPExcel_Cell::columnIndexFromString($highestColumm); //字母列转换为数字列 如:AA变为27
 
/** 循环读取每一行 */
for ($row = 1; $row <= $highestRow; $row++){//行数是以第1行开始
	//循环读取每一列
    for ($column = 0; $column < $highestColumm; $column++) {//列数是以第0列开始
        $columnName = PHPExcel_Cell::stringFromColumnIndex($column);
        echo $columnName.$row.":".$sheet->getCellByColumnAndRow($column, $row)->getValue()." ";
    }
    echo "<br>";
}