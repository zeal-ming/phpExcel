<?php
header("Content-Type:text/html;charset=utf-8");
$dir = dirname(__FILE__);
require_once $dir . '/PHPExcel.php';


$filename = './chart_example.xlsx';
// 全部加载excel文件(全部加载)
$objPHPExcel = PHPExcel_IOFactory::load($filename);

$sheetCount = $objPHPExcel->getSheetCount();
// for ($i=0; $i < $sheetCount; $i++) 
// { 
// 	//读取每个sheet
// 	$data = $objPHPExcel->getSheet($i)->toArray();
// 	echo "<pre>";
// 	print_r($data);
// }


//循环读取sheet
foreach ($objPHPExcel->getWorksheetIterator() as $sheet) {
	//循环读取每一行
	foreach ($sheet->getRowIterator() as $row) {
		if($row->getRowIndex() < 2)
		{
			continue;
		}
		//循环读取每个单元格
		foreach ($row->getCellIterator() as $cell) {
			$data = $cell->getValue();
			echo $data." ";
		}
		echo "<br>";
	}
}
