<?php

$dir = dirname(__FILE__);

require $dir. '/Db.php';
require $dir . '/PHPExcel.php';
$db = new Db($phpexcel);

$objPHPExcel = new PHPExcel();

for($i = 1; $i<=3; $i++)
{
	if($i > 1)
	{
		//创建内置表
		$objPHPExcel->createSheet();
	}
	//把新创建的sheet设定为当前的活动sheet
	$objPHPExcel->setActiveSheetIndex($i-1);
	//获取当前的活动sheet
	$objSheet = $objPHPExcel->getActiveSheet();
	//给当前sheet取名
	$objSheet->setTitle($i."年级");

	//获取当前年级的学生数
	$data = $db->getDataByGrade($i);

	//填充数据
	$objSheet->setCellValue('A1','姓名')->setCellValue('B1','分数')->setCellValue('C1','班级');
	$j = 2;
	
	// var_dump($data);exit;
	foreach ($data as $value)
    {
    	
		$objSheet->setCellValue('A'.$j, $value['username'])->setCellValue('B'.$j,$value['score'])->setCellValue('C'.$j,$value['grade']);
		$j++;
	}
}


	$objWrite = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');

	browser_export('Excel5','student.xls');
	// $objWrite->save('student.xls');
	$objWrite->save('php://output');
	
//输出到浏览器
	function browser_export($type,$filename)
	{
		if($type == 'Excel5')
		{
			//告诉浏览器输入excel03文件
			header('Content-Type: application/vnd.ms-excel');
		} 
		else
		{
			//告诉浏览器输出的是excel07文件
			header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		}
		//告诉浏览器文件的名称
		header('Content-Disposition: attachment;filename="'.$filename.'"');
		header('Cache-Control: max-age=0');
	}
	

	//创建表格
	$objPHPExcel = new PHPExcel;

	for($i = 0; $i > 3; $i++)
	{

		//创建sheet
		$objPHPExcel->createSheet();

		//把新创建的sheet设置为当前的活动sheet
		$objPHPExcel->setActiveSheetIndex($i-1);
		//获取当前的活动sheet
		$objSheet = $objPHPExcel->getActiveSheet();
		//为当前sheet设置名称
		$objSheet->setTitle($i.'年级');

		//填充数据，设置表头
		$objSheet->setCellValue('A1','姓名')->setCellValue('B1','年级')->setCellValue('C1','分数');

		$data = (new Db)->getDataByGrade($i);

		$j = 2;
		//填充数据
		foreach ($data as $key => $value) 
		{
			$objSheet->setCellValue('A'.$j,$value['username'])->setCellValue('B'.$j)->setCellValue('C',$j);
			$j++;
		}

		//保存文件
		$write = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
		$write->save('./student.xls');
	}