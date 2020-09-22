<?php

namespace App\Model;

use Illuminate\Database\Eloquent\Model;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class Export extends Model
{
    /**
     * 导出
     * @param $lists
     * @param $excelDir
     * @param $fileName
     * @return array
     */
    public function exportExcel($lists, $excelDir, $fileName)
    {
        $CpnList = new CpnList();

        $excelTitle = $CpnList->cpnList;

        $titleKey = array_keys($excelTitle);

        // excel处理
        $Spreadsheet = new Spreadsheet();

        $objSheet = $Spreadsheet->getActiveSheet();
        $objSheet->setTitle('清单');

        // 表头样式处理
        $Spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight('30');

        // 表头数据处理
        foreach ($titleKey as $i => $title) {
            $col = $this->IntToChr($i);

            // 居中样式
            $Spreadsheet->getActiveSheet()->getStyle($col . '1')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '1')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '1')->getAlignment()->setWrapText(true);

            // 赋值
            $Spreadsheet->getActiveSheet()->setCellValue($col . '1', @$excelTitle[$title][0]);
            $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
        }

        // 从0开始的id
        $index = 0;

        // 从第二行开始写数据
        $rows = 1;

        // 数据处理
        foreach ($lists as $k => $list) {
            $index += 1;
            $rows += 1;
            foreach ($titleKey as $i => $title) {
                $cols = $this->IntToChr($i);

                if ($cols == 'A') {
                    $data[$index][$cols . $rows] = $index;
                    $Spreadsheet->getActiveSheet()->setCellValue($cols . $rows, $index);
                } else {
                    $Spreadsheet->getActiveSheet()->setCellValue($cols . $rows, @$list[$title]);
                }
            }
        }

        // 保存文件
        try
        {
            $objWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($Spreadsheet, 'Xls');
            $filename_ = $excelDir . DIRECTORY_SEPARATOR . $fileName;
            $objWriter->save($filename_);

            return [
                'status'    => true,
                'file_path' => $filename_,
                'file_name' => $fileName
            ];
        } catch (\Exception $e)
        {
            return [
                'status'    => false,
                'file_path' => null,
                'file_name' => null
            ];
        }
    }

    /**
     * 数字转字母 （类似于Excel列标）
     * @param Int $index 索引值
     * @param Int $start 字母起始值
     * @return String 返回字母
     */
    function IntToChr($index, $start = 65)
    {
        $str = '';
        if (floor($index / 26) > 0) {
            $str .= $this->IntToChr(floor($index / 26) - 1);
        }
        return $str . chr($index % 26 + $start);
    }
}
