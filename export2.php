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


//填充数据
$gradeInfo =  $db->getAllGrades();
$index = 0;

foreach ($gradeInfo as $g_k => $g_v) {
	$classinfo = $db->getClassByGrade($g_v['grade']);

	//设置年段信息
	$gradeIndex = getCells($index * 2);
	
	$objSheet->setCellValue($gradeIndex.'2','高'.$g_v['grade']);
	foreach ($classinfo as $c_k => $c_v) 
	{
		$nameIndex = getCells($index*2);
		$scoreIndex = getCells($index*2+1);
		//在第三行设置班级信息
		$objSheet->setCellValue($nameIndex.'3',$c_v['class'].'班');
		//设置自动换行，注意写在换行符前面
		$objSheet->getStyle($nameIndex)->getAlignment()->setWrapText(true);
		//在第四行设置班级信息信息
		$objSheet->setCellValue($nameIndex.'4',"姓名\n换行")->setCellValue($scoreIndex.'4','分数');
		//合并班级的单元格
		$objSheet->mergeCells($nameIndex.'3'.':'.$scoreIndex.'3');
		//设置班级的背景颜色
		$objSheet->getStyle($nameIndex.'3'.':'.$scoreIndex.'3')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('52e351');
		//调整班级所在行文字大小
	$objSheet->getStyle($nameIndex.'3'.':'.$scoreIndex.'3')->getFont()->setSize(16)->setBold(true);

	//填充班级边框
	$classBorderStyle = getBorderStyle('e351ca');
	$objSheet->getStyle($nameIndex.'3'.':'.$scoreIndex.'3')->applyFromArray($classBorderStyle);
		//获取班级学生信息
		$info = $db->getDataByClassAndGrade($c_v['class'],$g_v['grade']);
		// var_dump($info);exit;
		$j = 5;

		foreach ($info as $key => $value) 
		{
			// $objSheet->setCellValue($nameIndex.$j,$value['username'])->
			// setCellValue($scoreIndex.$j,$value['score']);
			// 使用setValueExplicit设置单元格的显示类型，防止数字过长显示成科学计数法
			$objSheet->setCellValue($nameIndex.$j,$value['username'])->setCellValueExplicit($scoreIndex.$j,'23942342342323', PHPExcel_Cell_DataType::TYPE_STRING);
			$j++;
		}
		 $index++;
		 //合并单元格
		 $endIndex = getCells($index*2 -1);
		 $objSheet->mergeCells($gradeIndex.'2'.':'.$endIndex.'2');
	}

	//设置文字默认水平，垂直居中
	$objSheet->getDefaultStyle()->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	//设置默认字体大小
	$objSheet->getDefaultStyle()->getFont()->setName('微软雅黑')->setSize('14');
	//调整年级所在行文字样式
	$objSheet->getStyle($gradeIndex.'2'.':'.$endIndex.'2')->getFont()->setSize(20)->setBold(true);

	//设置年级颜色
	$objSheet->getStyle($gradeIndex.'2'.':'.$endIndex.'2')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('e36951');

	//填充年级边框
	$gradeBorderStyle = getBorderStyle('e3df51');
	$objSheet->getStyle($gradeIndex.'2'.':'.$endIndex.'2')->applyFromArray($gradeBorderStyle);

}
//保存数据
$writer = PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5');
$writer->save('./student/'.time().'.xls');

//填充数据
function fillDatas($db,$objSheet)
{

	$index = 0;
	//查询所有的年级
	$grades = $db->getAllGrades();
	foreach ($grades as $g_k => $g_v)
	{
		

		//获取该年级所有班级
		$classinfo = $db->getClassByGrade($g_v['grade']);

			foreach ($classinfo as $c_k => $c_v) 
			{
				if(getCells($index*2))
				{
					//获取姓名所在列
					$nameIndex = getCells($index*2);
					//获取分数所在列
					$scoreIndex = getCells($index*2 + 1);
				}
				
				// var_dump($scoreIndex);exit;
				// echo $nameIndex;
				//通过班级，获取班级所有学生信息
				$students = $db->getDataByClass($c_v['class']);
				//从第五行开始
				$row = 5;
				//循环填充数据
				foreach($students as $stu_v)
				{
					$nameIndex = $nameIndex.$row;
					$scoreIndex = $scoreIndex.$row;
					// var_dump($scoreIndex);exit
					$objSheet->setCellValue($nameIndex.$row, $stu_v['username'])->setCellValue($scoreIndex.$row, $stu_v['score']);
					$row++;
				}
				$index++;
			}
	}
}

function getCells($index)
{
	$arr = range('A','Z');
	if(count($arr) > $index)
	{
		return $arr[$index];
	}
	else
	{
		exit('数组越界');
	}
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