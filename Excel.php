<?php
/**
 * Created by PhpStorm.
 * User: 
 * Date: 
 * Time: 
 */

namespace lib;
require_once 'classes/PHPExcel.php';
use PHPExcel;
use PHPExcel_Reader_Excel2007;
use PHPExcel_Cell;
use PHPExcel_Writer_Excel2007;
class Excel
{
    /**
     * @param $filetmp_name  
     * @param array
     * @param array $title 
     * @return array
     */
    function import($filetmp_name,$zm = array(),$title = array()){
        header('content-type:text/html;charset=utf-8');
        $file=$filetmp_name;
        $PHPExcel  = new PHPExcel();
        $PHPReader = new PHPExcel_Reader_Excel2007();

        $PHPExcel = $PHPReader->load($file);

        $currentSheet = $PHPExcel->getSheet();

        $allColumn = PHPExcel_Cell::columnIndexFromString($currentSheet->getHighestColumn());

        $allRow = $currentSheet->getHighestRow();
        $insertData=array();
        for($i=2;$i<=$allRow;$i++){
            for($j=2;$j<=$allColumn;$j++){
                $value = $currentSheet->getCell($zm[$j-1].$i)->getValue();
                $insertData[$i][$title[$j-1]]=$value;
            }
        }

       $insertData=array_values($insertData);
       return $insertData;
    }

    //下载模板
    /**
     * @param array $zm  //['A','B','C']
     * @param array $title   标题
     */
    function demo($zm = array(),$title = array()){
        $objPHPExcel = new PHPExcel();
        $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
        $num=1;
        foreach($title as $key=>$val){
            $objPHPExcel->getActiveSheet()->setCellValue($zm[$key].$num, $val);
        }
        header("Content-Type:application/download");
        header('Content-Disposition:attachment;filename="export.xls"');
        $objWriter->save('php://output');
    }

    /**
     * @param $select
     * @param array $zm  ['A','B','C']
     * @param array $title    比如标题
     */

    function export($select,$zm = array(),$title = array()){
        $objPHPExcel = new PHPExcel();
        $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);

          $objPHPExcel->setActiveSheetIndex();
        $objactSheet = $objPHPExcel->getActiveSheet()->setTitle('user');
        $num=1;
        foreach($title as $key=>$val){
            $objactSheet->setCellValue($zm[$key].$num, $val);
        }

        foreach($select as $k=>$v){
            $num++;
            $k=0;
            foreach($v as $kk=>$vv){
                $objactSheet->setCellValue($zm[$k].$num, $vv);
                $k+=1;
            }
        }
        header("Content-Type:application/download");
        header('Content-Disposition:attachment;filename="export.xls"');
        $objWriter->save('php://output');
    }

}
