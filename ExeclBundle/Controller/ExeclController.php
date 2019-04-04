<?php

namespace ExeclBundle\Controller;

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Reader\Xls;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

class ExeclController
{
    /**
     * @param string $pfilename
     *               文件地址
     * @param int $sheet
     *            文件行数 默认从第一行取
     * @param int $columnCnt
     *            文件列 默认最大列
     * @throws \Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     *
     */
    public function ImportExecl($pfilename,$sheet=0,$columnCnt=0,$title)
    {
        $objRead = IOFactory::createReader('Xlsx');

        if (!$objRead->canRead($pfilename)) {

            $objRead = IOFactory::createReader('Xls');

            if (!$objRead->canRead($pfilename)) {
                throw new \Exception('只支持导入Excel文件！');
            }
        }
        empty($options) && $objRead->setReadDataOnly(true);

        $obj = $objRead->load($pfilename);

        $currSheet = $obj->getSheet($sheet);

        if (isset($options['mergeCells'])) {
            /* 读取合并行列 */
            $options['mergeCells'] = $currSheet->getMergeCells();
        }
        if ($columnCnt == 0) {
            /* 取得最大的列号 */
            $columnH = $currSheet->getHighestColumn();

            $columnCnt = Coordinate::columnIndexFromString($columnH);
        }
        $rowCnt = $currSheet->getHighestRow();

        $data = [];
        /* 读取内容 */
        for ($_row = 1; $_row <= $rowCnt; $_row++) {
            $isNull = true;
            for ($_column = 1; $_column <= $columnCnt; $_column++) {
                $cellName = Coordinate::stringFromColumnIndex($_column);
                $cellId = $cellName . $_row;
                $cell = $currSheet->getCell($cellId);
                if (isset($options['format'])) {
                    /* 获取格式 */
                    $format = $cell->getStyle()->getNumberFormat()->getFormatCode();
                    /* 记录格式 */
                    $options['format'][$_row][$cellName] = $format;
                }
                if (isset($options['formula'])) {
                    /* 获取公式，公式均为=号开头数据 */
                    $formula = $currSheet->getCell($cellId)->getValue();
                    if (0 === strpos($formula, '=')) {
                        $options['formula'][$cellName . $_row] = $formula;
                    }
                }
                if (isset($format) && 'm/d/yyyy' == $format) {
                    /* 日期格式翻转处理 */
                    $cell->getStyle()->getNumberFormat()->setFormatCode('yyyy/mm/dd');
                }
                $data[$_row][$cellName] = trim($currSheet->getCell($cellId)->getFormattedValue());
                if (!empty($data[$_row][$cellName])) {
                    $isNull = false;
                }
            }

            if ($isNull) {
                unset($data[$_row]);
            }
        }

        if(array_values(array_shift($data)) != explode(',',$title)){
            throw new \Exception('请按模板文件格式上传！');
        }
        return $data;
    }


    /**
     * @param array $title 例如：[姓名,性别]
     * @param array $data  例如：[['nickname'=>'张三','sex'=>'男'],
     *                           ['nickname'=>'李四','sex'=>'女']]
     * @param string $fileName
     * @return bool
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     *
     */
    public function exportExecl($title,$data,$fileName)
    {
        if(empty($data)){
            return false;
        }
        set_time_limit(0);
        $arrData = array_merge(array($title), $data);
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $spreadsheet->getActiveSheet()->fromArray($arrData);
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        header('Content-Description: File Transfer');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename='.$fileName.'.xlsx');
        header('Cache-Control: max-age=0');
        $writer->save('php://output');
        $fp = fopen('php://output', 'a');
        mb_convert_variables('GBK', 'UTF-8', $columns);
        fputcsv($fp, $columns);
        $dataNum = count( $arrData );
        $perSize = 1000;
        $pages = ceil($dataNum / $perSize);
        for ($i = 1; $i <= $pages; $i++) {
            foreach ($arrData as $item) {
                mb_convert_variables('GBK', 'UTF-8', $item);
                fputcsv($fp, $item);
            }
            ob_flush();
            flush();
        }
        fclose($fp);
        return true;
    }
}
