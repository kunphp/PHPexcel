<?php 
include 'pdo.php';
/** 
 * 数据导出 
 * @param array $title   标题行名称 
 * @param array $data   导出数据 
 * @param string $fileName 文件名 
 * @param string $savePath 保存路径 
 * @param $type   是否下载  false--保存   true--下载 
 * @return string   返回文件全路径 
 * @throws PHPExcel_Exception 
 * @throws PHPExcel_Reader_Exception 
 */  
function exportExcel($title=array(), $data=array(), $fileName='', $savePath='./', $isDown=true){
   include('Classes/PHPExcel.php');
    $obj = new PHPExcel();  
  
    //横向单元格标识  
    $cellName = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ');  
      
    $obj->getActiveSheet(0)->setTitle('项目列表');   //设置sheet名称  
    $_row = 1;   //设置纵向单元格标识  
    if($title){  
        $_cnt = count($title);  
        $obj->getActiveSheet(0)->mergeCells('A'.$_row.':'.$cellName[$_cnt-1].$_row);   //合并单元格  
        $obj->setActiveSheetIndex(0)->setCellValue('A'.$_row, '数据导出：'.date('Y-m-d H:i:s'));  //设置合并后的单元格内容  
        $_row++;  
        $i = 0;  
        foreach($title AS $v){   //设置列标题  
            $obj->setActiveSheetIndex(0)->setCellValue($cellName[$i].$_row, $v);  
            $i++;  
        }  
        $_row++;  
    }  
  
    //填写数据  
    if($data){  
        $i = 0;  
        foreach($data AS $_v){  
            $j = 0;  
            foreach($_v AS $_cell){  
                $obj->getActiveSheet(0)->setCellValue($cellName[$j] . ($i+$_row), $_cell);  
                $j++;  
            }  
            $i++;  
        }  
    }  
      
    //文件名处理  
    if(!$fileName){  
        $fileName = uniqid(time(),true);  
    }  
  
    $objWrite = PHPExcel_IOFactory::createWriter($obj, 'Excel2007');  
  
    if($isDown){   //网页下载  
        header('pragma:public');  
        header("Content-Disposition:attachment;filename=$fileName.xls");  
        $objWrite->save('php://output');exit;  
    }  
  
    $_fileName = iconv("utf-8", "gb2312", $fileName);   //转码  
    $_savePath = $savePath.$_fileName.'.xlsx';  
     $objWrite->save($_savePath);  
  
   return $savePath.$fileName.'.xlsx';  
}  
  
//默认主数据库IP地址
$systemConfig['DB_HOST'] = '';

//默认主数据库名称
$systemConfig['DB_NAME'] = '';

//默认主数据库帐号
$systemConfig['DB_USER'] = '';

//默认主数据库密码
$systemConfig['DB_PWD'] = '';

//默认主数据库端口
$systemConfig['DB_PORT'] = 3306;

//默认主数据库编码
$systemConfig['DB_CHARSET'] = 'UTF8';

//默认主数据库其扩展配置
$systemConfig['DB_OPTIONS'] = array(
    //如果为true => 数据库连接为持久化连接
    PDO::ATTR_PERSISTENT => false,

    //错误模式 :
    //PDO::ERRMODE_SILENT       =>    默认模式，只简单地设置错误码，生产环境推荐
    //PDO::ERRMODE_WARNING      =>    引发 E_WARNING 错误；
    //PDO::ERRMODE_EXCEPTION    =>    抛出 exceptions 异常 (开发模式推荐)
    PDO::ATTR_ERRMODE => PDO::ERRMODE_SILENT
);
$pdo=new pdoModel($systemConfig);
$searchSql = "";
//要导入的数据
$data=$pdo->getAll($searchSql);
//字段对应的名 
$title=['','','','','',''];
//文件名
$fileName='';
exportExcel($title,$data,$fileName);
?>