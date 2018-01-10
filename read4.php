<?php

error_reporting(E_ALL);
 
date_default_timezone_set('Asia/ShangHai');
 
/** PHPExcel_IOFactory */
require_once './PHPExcel.php';
 
 $filename = "student/student.xls";
 //自动获取文件的类型
$fileType = PHPExcel_IOFactory::identify($filename);

// Check prerequisites
if (!file_exists($filename)) {
	exit("not found $filename.\n");
}

$reader = PHPExcel_IOFactory::createReader($fileType); //设置以Excel5格式(Excel97-2003工作簿)
$PHPExcel = $reader->load($filename); // 载入excel文件
$sheet = $PHPExcel->getSheet(0); // 读取第一個工作表
$highestRow = $sheet->getHighestRow(); // 取得总行数
$highestColumm = $sheet->getHighestColumn(); // 取得总列数
 
/** 循环读取每个单元格的数据 */
for ($row = 1; $row <= $highestRow; $row++){//行数是以第1行开始
    for ($column = 'A'; $column <= $highestColumm; $column++) {//列数是以A列开始
        $dataset[] = $sheet->getCell($column.$row)->getValue();
        echo $column.$row.":".$sheet->getCell($column.$row)->getValue()."<br />";
    }
}