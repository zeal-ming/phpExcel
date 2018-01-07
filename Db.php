<?php

require dirname(__FILE__).'/dbconfig.php';

class Db
{
	private $conn = null;

	public function __construct($config)
	{
	//连接数据库
		$this->conn = mysql_connect($config['host'],$config['username'],$config['password']) or die(mysql_error());

	//选择数据库
	mysql_select_db($config['database'], $this->conn) or die(mysql_error());

	//设置mysql编码
	mysql_query("set names ".$config['charset']) or die(mysql_error());

	}

	/*
	根据传入的sql语句，返回结果集

	 */
	public static function getResult($sql)
	{
		$resource = mysql_query($sql) or die(mysql_error());

		$res = array();
		while(($row = mysql_fetch_assoc($resource)))
		{
			$res[] = $row;
		}

		return $res;
	}

	//获取所有年级
	public function getAllGrades()
	{
		
		$sql = "select grade from student group by grade";

		$res = self::getResult($sql);
		return $res;
	}
	//通过年级获取班级
	public function getClassByGrade($grade)
	{	
		$grade = intval($grade);
		$sql = "select class from student where grade=$grade order by score desc";
		$res = self::getResult($sql);
		return $res;
	}
	//通过班级获取所有学生信息
	public function getDataByClassAndGrade($class,$grade)
	{
		$class = intval($class);
		$grade = intval($grade);
		$sql = "select score,username from student where grade=$grade and class=$class order by class desc";
		// echo $sql;
		$res = self::getResult($sql);
		return $res;
	}



}