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

//填充数据
$objSheet->fromArray($array);

//1右侧显示
$labels = [
	new PHPExcel_Chart_DataSeriesValues('String','Worksheet!$B$1',null,1),
	new PHPExcel_Chart_DataSeriesValues('String','Worksheet!$C$1',null,1),
	new PHPExcel_Chart_DataSeriesValues('String','Worksheet!$D$1',null,1),
];

//2底部显示刻度
$xLabels = [
	new PHPExcel_Chart_DataSeriesValues('String','Worksheet!$A$2:$A$4',null,3),
];

//3.取得绘图所需数据
$datas = [
	new PHPExcel_Chart_DataSeriesValues('Number','Worksheet!$B$2:$B$4',null,3),
	new PHPExcel_Chart_DataSeriesValues('Number','Worksheet!$C$2:$C$4',null,3),
	new PHPExcel_Chart_DataSeriesValues('Number','Worksheet!$D$2:$D$4',null,3),
];

//4.制作框架
$series = [
	new PHPExcel_Chart_DataSeries(
		PHPExcel_Chart_DataSeries::TYPE_LINECHART,
		PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
		range(0, count($labels)-1),
		$labels,
		$xLabels,
		$datas
		),
];

//5.图片节点中显示数据
$layout = new PHPExcel_Chart_Layout();
$layout->setShowVal(true);

//6.生成图表
$areas = new PHPExcel_Chart_PlotArea($layout,$series);
$title = new PHPExcel_Chart_Title('学生成绩表');
$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT,$layout,false);
$chart = new PHPExcel_Chart(
	  'line_chart',
	  $title,
	  $legend,
	  $areas,
	  true,
	  false,
	  null
	);
//7.给图表定位
$chart->setTopLeftPosition('A7')->setBottomRightPosition('K25');
//8.把图表添加到表格中
$objSheet->addChart($chart);

//保存数据
$write = PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel2007');
$write->setIncludeCharts(true);
$write->save('./zm.xls');
