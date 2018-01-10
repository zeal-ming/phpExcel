<?php

$dir = dirname(__FILE__);

require $dir. '/Db.php';
require $dir . '/PHPExcel.php';
$db = new Db($phpexcel);

$objPHPExcel = new PHPExcel();
$objSheet = $objPHPExcel->getActiveSheet();

//准备数据
$array = [
	['标准','一班','二班','三班'],
	['不及格','33','44','55'],
	['良好','66','67','77'],
	['优秀','86','90','100'],
];
 
//直接填充到单元格中
$objSheet->fromArray($array);

//开始图标代码编写
//右侧显示的标签，取得绘制图标的标签
$labels = array(
	new PHPExcel_Chart_DataSeriesValues('String','Worksheet!$B$1',null,1),
	new PHPExcel_Chart_DataSeriesValues('String','Worksheet!$C$1',null,1),
	new PHPExcel_Chart_DataSeriesValues('String','Worksheet!$D$1',null,1),
	);
//底部显示的刻度，取得图标x周的刻度
$xLabels = array(
	new PHPExcel_Chart_DataSeriesValues('String','Worksheet!$A$2:$A$4',null,3),
	);

//取得绘制所需要数据
$datas = [
	new PHPExcel_Chart_DataSeriesValues('Number','Worksheet!$B$2:$B$4',null,3),
	new PHPExcel_Chart_DataSeriesValues('Number','Worksheet!$C$2:$C$4',null,3),
	new PHPExcel_Chart_DataSeriesValues('Number','Worksheet!$D$2:$D$4',null,3),
];

//根据取得的东西做出一个图表的框架
$series = [
	new PHPExcel_Chart_DataSeries(
		PHPExcel_Chart_DataSeries::TYPE_LINECHART,
		PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
		range(0,count($labels)-1),
		$labels,
		$xLabels,
		$datas
		),
	];

//在图形节点中显示数据
$layout = new PHPExcel_Chart_Layout();
$layout->setShowVal($true);

//生成图表
$areas = new PHPExcel_Chart_PlotArea($layout,$series);
$title = new PHPExcel_Chart_Title('高一学生成绩分布');
$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT,$layout,false);
$yTitle =  new PHPExcel_Chart_Title("value(人数)");

$chart = new PHPExcel_Chart(
	'line_chart',
	$title,
	// $legend,
	null,
	$areas,
	true,
	0,
	null,
	$yTitle
	);

//给定图表所在的表格位置
$chart->setTopLeftPosition('A7')->setBottomRightPosition('K25');
//把chart添加到表格中
$objSheet->addChart($chart);
$objWrite = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWrite->setIncludeCharts(true);
$objWrite->save('./student/'.time().'.xls');
	
