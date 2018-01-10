<?php
header("Content-Type:text/html;charset=utf-8");
$dir = dirname(__FILE__);
require_once $dir . '/PHPExcel.php';

$filename = './student/student.xls';

//自动获取文件的类型
$fileType = PHPExcel_IOFactory::identify($filename);
//选择加载,获取了文件读取操作对象
$objReader = PHPExcel_IOFactory::createReader($fileType);
$sheetName = 'Sheet1';
//加载某个sheet
 $objReader->setLoadSheetsOnly($sheetName);
$objPHPExcel = $objReader->load($filename);
//循环读取sheet
foreach ($objPHPExcel->getWorksheetIterator() as $sheet) {
	//循环读取每一行
	foreach ($sheet->getRowIterator() as $row) {

		//循环读取每个单元格
		foreach ($row->getCellIterator() as $cell) {
			$data = $cell->getValue();
			echo $data." ";
		}
		echo "<br>";
	}
}