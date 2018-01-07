<?php

require('./PHPExcel.php');
// require('./PHPExcel/IOFactory.php');
// require('./PHPExcel/Worksheet.php');

/*
1.新建excel表格
2.创建sheet内置表
3.填充数据
4.保存数据
 */
//相当于在桌面创建一个excel表格
$excel = new PHPExcel;

//获取当前活动sheet的操作对象
$actSheet = $excel->getActiveSheet();
//给当前活动设置名称
$actSheet->setTitle('demo');
//填充数据
$actSheet->setCellValue('A1','张铭')->setCellValue('B1','分数');
//加载数据块
$array = array(
	array(' ','姓名','性别'),
	array(' ','张铭', '男'),
	);
$actSheet->fromArray($array);
// $actSheet->setmergeCeils('A1:B1');

$objWrite = PHPExcel_IOFactory::createWriter($excel,'Excel2007');
$objWrite->save("./demo1.xlsx");