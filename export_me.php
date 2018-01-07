<?php

require './Db.php';
require './PHPExcel.php';

//获取年级
$db = new Db($phpexcel);

// var_dump($grades);exit;
//创建表格
$objPHPExcel = new PHPExcel;

//获取当前活动sheet
$objSheet = $objPHPExcel->getActiveSheet();
//设置sheet名称
$objSheet->setTitle('班级成绩');
//设置所有文字默认大小
$objSheet->getDefaultStyle()->getFont()->setName('微乳雅黑')->setSize('14');
//设置所有文字居中方式
$objSheet->getDefaultStyle()->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

//填充数据
//获取所有年级
$gradeInfo =  $db->getAllGrades();
$index = 0;

foreach ($gradeInfo as $g_k => $g_v)
{
	//根据年级获取所有班级
	$classinfo = $db->getClassByGrade($g_v['grade']);
	//获取年级所在列
	$gradeIndex = getCells($index*2);

	$objSheet->mergeCells();
	foreach ($classinfo as $c_k => $c_v) 
	{
		//根据年级，班级获取所有学生信息
		$students = $db->getDataByClassAndGrade($c_v['class'], $g_v['grade']);
		// var_dump($students);exit;
		//获取姓名所在列
		$nameIndex = getCells($index*2);
		//获取分数所在列
		$scoreIndex = getCells($index*2 + 1);

		//根据所在列，填充姓名，成绩
		$objSheet->setCellValue($nameIndex.'4','姓名')->setCellValue($scoreIndex.'4', '成绩');
		$objSheet->mergeCells($nameIndex.'3'.':'.$scoreIndex.'3');
		//填充班级名称
		$objSheet->setCellValue($nameIndex.'3',$c_v['class'].'班');
		//设置字体大小
		$objSheet->getStyle($nameIndex.'3'.':'.$scoreIndex.'3')->getFont()->setSize('16')->setBold(true);

		//设置从第四行开始填充
		$row = 4;
		foreach ($students as $key => $value) 
		{
			//填充学生信息
			$objSheet->setCellValue($nameIndex.$row, $value['username'])->setCellValue($scoreIndex.$row, $value['score']);
			$row++;
		}
		$index++;
		
	}

	$endIndex = getCells($index*2 - 1);
	//填充字体
	$objSheet->setCellValue($gradeIndex.'2','高'.$g_v['grade']);
	//合并年级
	$objSheet->mergeCells($gradeIndex.'2'.':'.$endIndex.'2');
	//给年级设置背景颜色
	$objSheet->getStyle($gradeIndex.'2'.':'.$endIndex.'2')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('e36951');
	//给年级设置字体大小
	$objSheet->getStyle($gradeIndex.'2'.':'.$endIndex.'2')->getFont()->setSize('20')->setBold(true);
	//给年级设置边框
	$gradeBorderStyle = getBorderStyle('e3df51');
	$objSheet->getStyle($gradeIndex.'2'.':'.$endIndex.'2')->applyFromArray($gradeBorderStyle);

}

//保存数据
$writer = PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5');
$writer->save('./student/'.time().'.xls');

//根据索引获取列
function getCells($index)
{
	$arr = range('A','Z');
	return $arr[$index];
}

//获取边框样式
function getBorderStyle($color)
{
	$styleArray = array(
	'borders' => array(
		'outline' => array(
			'style' => PHPExcel_Style_Border::BORDER_THICK,
			'color' => array('rgb' => $color),
			),
		),
	);

	return $styleArray;
}