<?php

namespace App\Model;

use function foo\func;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\SimpleType\JcTable;
use ZanySoft\Zip\Zip;

/**
 * 意见审核跟踪表
 * Class ProReviewTask
 * @package App\Model
 */
class ProNew extends Model
{
    /**
     * 获取元器件导出的数据
     * @param $main_task_id
     * @return array
     */
    public function getExportData($main_task_id)
    {
        $CpnFiles  = new CpnFiles();
        $listDatas = $CpnFiles->getListIds($main_task_id, 2);

        $cpnData = CpnImport::whereIn('list_id', array_column($listDatas,'id'))
            ->whereIn('result_pc', [2, 3, 4])
            ->with('AuxiliaryResultReplace')
            ->get()
            ->groupBy(['cpn_specification_model'])
            ->toArray();

        $codeData = CpnImport::whereIn('list_id', array_column($listDatas,'id'))
            ->whereIn('result_pc', [2, 3, 4])
            ->select('cpn_category_code')
            ->get()
            ->groupBy('cpn_category_code')
            ->toArray();

        $modelData = CpnFiles::whereIn('id', array_column($listDatas,'id'))
            ->get()
            ->toArray();

        $modelStructure = ModelStructure::where('main_task_id', $main_task_id)
            ->select('name', 'id', 'parent_id')
            ->get()
            ->toArray();

        $treePathArr = $this->getTreePath($modelStructure, $modelData);

        // 辅助审查结果页面 映射用
        $pc_arr = [
            1 => '纳入进口清单',
            2 => '纳入国产清单',
            3 => '不允许选用',
            4 => '进入用研结合审查',
        ];

        // 分批插入元器件数据
        foreach ($cpnData as $cpnCollectKey => $cpnCollect)
        {
            $cpnCollect = array_values($cpnCollect);
            $modelCount = count($cpnCollect);

            $list_id_collect = array_column($cpnCollect, 'list_id');
            $list_id_collect = array_unique($list_id_collect);

            // 装机位置
            $path                                     = $this->getAllPath($list_id_collect, $treePathArr);
            $cpnData[$cpnCollectKey][0]['equip_name'] = $path;
            $cpnData[$cpnCollectKey][0]['cpn_is_core_important'] == 1 ? '是' : '否';
            $cpnData[$cpnCollectKey][0]['result_pc'] = $pc_arr[$cpnData[$cpnCollectKey][0]['result_pc']];
            $cpnData[$cpnCollectKey][0]['hash_code'] = md5(time() . rand(0, 9999) . $cpnCollect[0]['id']);

            // 当这个规格型号的下的数据为多条的时候进行处理
            if ($modelCount > 1)
            {
                $cpn_category_code_collect = array_column($cpnCollect, 'cpn_category_code');
                $cpn_category_code_collect = array_unique($cpn_category_code_collect);
                $cpn_category_code_count   = count($cpn_category_code_collect);

                $cpn_name_collect = array_column($cpnCollect, 'cpn_name');
                $cpn_name_collect = array_unique($cpn_name_collect);
                $cpn_name_count   = count($cpn_name_collect);

                $cpn_country_collect = array_column($cpnCollect, 'cpn_country');
                $cpn_country_collect = array_unique($cpn_country_collect);
                $cpn_country_count   = count($cpn_country_collect);

                $cpn_quality_collect = array_column($cpnCollect, 'cpn_quality');
                $cpn_quality_collect = array_unique($cpn_quality_collect);
                $cpn_quality_count   = count($cpn_quality_collect);

                $cpn_package_collect = array_column($cpnCollect, 'cpn_package');
                $cpn_package_collect = array_unique($cpn_package_collect);
                $cpn_package_count   = count($cpn_package_collect);

                $necessity_collect = array_column($cpnCollect, 'necessity');
                $necessity_collect = array_unique($necessity_collect);
                $necessity_count   = count($necessity_collect);

                $safe_color_collect = array_column($cpnCollect, 'safe_color');
                $safe_color_collect = array_unique($safe_color_collect);
                $safe_color_count   = count($safe_color_collect);

                $proposed_safe_color_collect = array_column($cpnCollect, 'proposed_safe_color');
                $proposed_safe_color_collect = array_unique($proposed_safe_color_collect);
                $proposed_safe_color_count   = count($proposed_safe_color_collect);

                $cpn_is_core_important_collect = array_column($cpnCollect, 'cpn_is_core_important');
                $cpn_is_core_important_collect = array_unique($cpn_is_core_important_collect);
                $cpn_is_core_important_count   = count($cpn_is_core_important_collect);

                $cpn_period_collect = array_column($cpnCollect, 'cpn_period');
                $cpn_period_collect = array_unique($cpn_period_collect);
                $cpn_period_count   = count($cpn_period_collect);

                $cpn_ref_price_collect = array_column($cpnCollect, 'cpn_ref_price');
                $cpn_ref_price_collect = array_unique($cpn_ref_price_collect);
                $cpn_ref_price_count   = count($cpn_ref_price_collect);

                // 元器件类别
                if ($cpn_category_code_count > 1)
                {
                    $codeEnd                                         = $this->getCodeEnd($cpn_category_code_collect, $codeData);
                    $cpnData[$cpnCollectKey][0]['cpn_category_code'] = $codeEnd;
                }

                // 元器件名称
                if ($cpn_name_count > 1)
                {
                    $cpnNameEnd                             = implode(',', $cpn_name_collect);
                    $cpnData[$cpnCollectKey][0]['cpn_name'] = $cpnNameEnd;
                }

                // 国别地区
                if ($cpn_country_count > 1)
                {
                    $cpnNameEnd                                = implode(',', $cpn_country_collect);
                    $cpnData[$cpnCollectKey][0]['cpn_country'] = $cpnNameEnd;
                }

                // 质量等级
                if ($cpn_quality_count > 1)
                {
                    $cpnNameEnd                                = implode(',', $cpn_quality_collect);
                    $cpnData[$cpnCollectKey][0]['cpn_quality'] = $cpnNameEnd;
                }

                // 封装形式
                if ($cpn_package_count > 1)
                {
                    $cpnNameEnd                                = implode(',', $cpn_package_collect);
                    $cpnData[$cpnCollectKey][0]['cpn_package'] = $cpnNameEnd;
                }

                // 必要性
                if ($necessity_count > 1)
                {
                    $cpnNameEnd                              = implode(',', $necessity_collect);
                    $cpnData[$cpnCollectKey][0]['necessity'] = $cpnNameEnd;
                }

                // 安全等级颜色
                if ($safe_color_count > 1)
                {
                    $result                                   = empty($cpnData[$cpnCollectKey][0]['yield_safe_color']) ? $cpnData[$cpnCollectKey][0]['safe_color'] : $cpnData[$cpnCollectKey][0]['yield_safe_color'];
                    $cpnData[$cpnCollectKey][0]['safe_color'] = $result;
                }

                // 建议安全等级颜色
                if ($proposed_safe_color_count > 1)
                {
                    $result                                            = empty($cpnData[$cpnCollectKey][0]['yield_proposed_safe_color']) ? $cpnData[$cpnCollectKey][0]['proposed_safe_color'] : $cpnData[$cpnCollectKey][0]['yield_proposed_safe_color'];
                    $cpnData[$cpnCollectKey][0]['proposed_safe_color'] = $result;
                }

                // 是否核心关键器件
                if ($cpn_is_core_important_count > 1)
                {
                    $result                                              = empty($cpnData[$cpnCollectKey][0]['yield_is_core_important']) ? $cpnData[$cpnCollectKey][0]['cpn_is_core_important'] : $cpnData[$cpnCollectKey][0]['yield_is_core_important'];
                    $cpnData[$cpnCollectKey][0]['cpn_is_core_important'] = $result;
                }

                // 参考价格
                if ($cpn_period_count > 1)
                {
                    $range                                          = $this->getRange($cpn_period_collect);
                    $cpnData[$cpnCollectKey][0]['cpn_period_count'] = $range;
                }

                // 供货周期
                if ($cpn_ref_price_count > 1)
                {
                    $range                                       = $this->getRange($cpn_ref_price_collect);
                    $cpnData[$cpnCollectKey][0]['cpn_ref_price'] = $range;
                }

                // 计算机国产化替代审查
                if (!empty($cpnCollect[0]['auxiliary_result_replace']))
                {
                    $result                                                   = $this->getAuxList($cpnCollect);
                    $cpnData[$cpnCollectKey][0]['kb_aux_substitution_plan']   = $result['plan'];
                    $cpnData[$cpnCollectKey][0]['kb_aux_substitution_model']  = $result['model'];
                    $cpnData[$cpnCollectKey][0]['kb_aux_substitution_status'] = $result['status'];
                    $cpnData[$cpnCollectKey][0]['kb_aux_substitution_msg']    = $result['msg'];
                }
            }
        }

        return $cpnData;
    }

    /**
     * 获取每个list_id对应的节点位置
     * @param $array
     * @param $model
     * @return array|null
     */
    public function getTreePath($array, $model)
    {
        if (empty($array))
        {
            return null;
        }

        $items = array();
        foreach ($array as $value)
        {
            $items[$value['id']] = $value;
        }

        $return = [];
        foreach ($model as $mKey => $mValue)
        {
            $treePath = '';
            foreach ($items as $key => $value)
            {
                if (isset($items[$value['parent_id']]))
                {
                    $treePath .= '/' . $value['name'];
                } else
                {
                    $treePath .= $value['name'];
                }
            }
            $return[$mValue['id']] = $treePath;
        }

        return $return;
    }

    /**
     * 获取完整的路径
     * @param $arr1
     * @param $arr2
     * @return string
     */
    public function getAllPath($arr1, $arr2)
    {
        $return = '';

        foreach ($arr1 as $list_id)
        {
            $return .= $arr2[$list_id] . '；';
        }

        return $return;
    }

    /**
     * 元器件类别处理
     * @param $arr
     * @param $codeData
     * @return mixed
     */
    public function getCodeEnd($arr, $codeData)
    {
        $maxArr       = [];
        $codeArr      = [];
        $codeCountArr = [];
        $endArr       = [];

        // 获取数组长度
        foreach ($arr as $key => $value)
        {
            $maxArr[$key] = strlen($value);
        }

        // 获取最长长度
        $max = max($maxArr);

        // 循环将最长的放到$maxArr数组中去
        foreach ($maxArr as $key => $value)
        {
            $ArrKey = array_search($max, $maxArr);

            if ($ArrKey === false)
                break;

            unset($maxArr[$ArrKey]);

            $codeArr[$key] = $arr[$ArrKey];
        }

        // 当只有一个最长的时候直接返回
        if (count($codeArr) == 1)
        {
            return current($codeArr);
        }

        // 判断两个最长的那个使用的数量比较多
        foreach ($codeArr as $key => $value)
        {
            $codeCountArr[$key] = count($codeData[$value]);
        }

        // 当他们的使用数量都一样的情况下，将第一个返回
        if (count(array_unique($codeCountArr)) == 1)
        {
            return current($codeArr);
        }

        // 取使用量最多的数组
        $max = max($codeCountArr);

        // 将最多的数组取出
        foreach ($codeCountArr as $key => $value)
        {
            $ArrKey = array_search($max, $codeCountArr);

            if ($ArrKey === false)
                break;

            unset($codeCountArr[$ArrKey]);

            $endArr[$key] = $codeArr[$ArrKey];
        }

        // 返回最多使用量的情况
        return current($endArr);

    }

    /**
     * 获取参考价格和供货周期的范围
     * @param $arr
     * @return string
     */
    public function getRange($arr)
    {
        $max = max($arr);
        $min = min($arr);

        return $min . '-' . $max;
    }

    /**
     * 处理辅助审查的数据
     * @param $cpnCollect
     * @return array
     */
    public function getAuxList($cpnCollect)
    {
        $kb_aux_substitution_model  = '';
        $kb_aux_substitution_plan   = 0;
        $kb_aux_substitution_status = '';
        $kb_aux_substitution_msg    = '';

        foreach ($cpnCollect[0]['auxiliary_result_replace'] as $aux_key => $aux_result)
        {
            $kb_aux_substitution_model .= $aux_result['replace_cpn_specification_model'] . '[' . $aux_result['replace_cpn_manufacturer'] . '|' . $aux_result['replace_cpn_quality'] . '];';
            if (!empty($aux_result['experience']))
                $kb_aux_substitution_status .= $aux_result['replace_cpn_specification_model'] . '在' . $aux_result['experience'] . ';';
            if (!empty($aux_result['history']))
                $kb_aux_substitution_msg .= $aux_result['history'];
        }

        if ($kb_aux_substitution_plan == 1)
            $kb_aux_substitution_plan = '国产手册';
        elseif ($kb_aux_substitution_plan == 2)
            $kb_aux_substitution_plan = '推荐';
        else
            $kb_aux_substitution_plan = '';

        return [
            'model'  => $kb_aux_substitution_model,
            'plan'   => $kb_aux_substitution_plan,
            'status' => $kb_aux_substitution_status,
            'msg'    => $kb_aux_substitution_msg,
        ];
    }

    /**
     * 获取要插入的数据
     * @param $main_task_id
     * @return array
     */
    public function getInstallData($main_task_id)
    {
        $CpnFiles  = new CpnFiles();
        $listDatas = $CpnFiles->getListIds($main_task_id, 2);

        $cpnData = CpnImport::whereIn('list_id', array_column($listDatas,'id'))
            ->whereIn('result_pc', [2, 3, 4])
            ->with('AuxiliaryResultReplace')
            ->get()
            ->groupBy(['cpn_specification_model'])
            ->toArray();

        $installArr = [];
        foreach ($cpnData as $key => $value)
        {
            $rowArr = current($value);

            $row['main_task_id']            = $main_task_id;
            $row['list_id']                 = $rowArr['list_id'];
            $row['cpn_id']                  = $rowArr['id'];
            $row['cpn_specification_model'] = $rowArr['cpn_specification_model'];
            $row['cpn_manufacturer']        = $rowArr['cpn_manufacturer'];
            $row['hash_code']               = md5(time() . rand(0, 9999) . $rowArr['id']);

            array_push($installArr, $row);
        }

        return $installArr;
    }

    /**
     * 导出专家任务excel
     * @param        $lists
     * @param string $excelDir
     * @param string $fileName
     * @return array
     * @throws \PHPExcel_Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function export($lists, $excelDir, $fileName)
    {

        $excelStatus    = $this->exportExcel($lists, $excelDir, $fileName);
        $excelExtStatus = $this->exportExcelExt($excelDir, $fileName);
        $wordStatus     = $this->exportWord($lists, $excelDir, $fileName);

        if (!array_key_exists('status', $excelStatus))
            return [
                'status' => false,
                'msg'    => 'export 001',
            ];

        if (!array_key_exists('status', $excelExtStatus))
            return [
                'status' => false,
                'msg'    => 'export 002',
            ];

        if (!array_key_exists('status', $wordStatus))
            return [
                'status' => false,
                'msg'    => 'export 003',
            ];

        $zipPath = 'attachment/package/' . date('Y-m-d') . '/' . $fileName . "数据包" . date('YmdHis') . ".zip";

        $ModelStructure = new ModelStructure();
        $res            = $ModelStructure->zipPackage($excelDir, $zipPath);

        if ($res === true)
        {
            return [
                'status'    => $res,
                'file_path' => $zipPath,
                'file_name' => $zipPath
            ];
        }

        return [
            'status'    => $res,
            'file_path' => $zipPath,
            'file_name' => $zipPath
        ];


    }

    /**
     * 导出审查清单01
     * @param $lists
     * @param $excelDir
     * @param $fileName
     * @return array
     * @throws \PHPExcel_Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function exportExcel($lists, $excelDir, $fileName)
    {
        $excel = new Excel();

        $excelTitle = $excel->importTitle;

        $titleKey = array_keys($excelTitle);

        // excel处理
        $Spreadsheet = new Spreadsheet();

        $objSheet = $Spreadsheet->getActiveSheet();
        $objSheet->setTitle('清单');

        // 表头
        $Spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight('30');
        $Spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight('40');

        $chunkList = count($lists) + 3;

        foreach ($titleKey as $i => $title)
        {
            $col = $excel->IntToChr($i);

            $Spreadsheet->getActiveSheet()->setCellValue($col . '2', @$excelTitle[$title][0]);

            $Spreadsheet->getActiveSheet()->getStyle($col . '2')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '2')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '2')->getAlignment()->setWrapText(true);

            $Spreadsheet->getActiveSheet()->getStyle($col . '1')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '1')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '1')->getAlignment()->setWrapText(true);

            // 设置列宽,默认10
            switch ($col)
            {
                case 'A':
                    $objSheet->mergeCells('A1:A2');
                    $Spreadsheet->getActiveSheet()->setCellValue($col . '1', '序号');
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(10);
                    break;
                case 'B':
                    $objSheet->mergeCells('B1:B2');
                    $Spreadsheet->getActiveSheet()->setCellValue($col . '1', '装机位置');
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(60);
                    break;
                case 'C':
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(30);
                    break;
                case 'D':
                    $objSheet->mergeCells('D1:M1');
                    $Spreadsheet->getActiveSheet()->setCellValue($col . '1', '进口电子元器件信息');
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(35);
                    break;
                case 'E':
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(35);
                    break;
                case 'F':
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(35);
                    break;
                case 'N':
                    $objSheet->mergeCells('N1:N2');
                    $Spreadsheet->getActiveSheet()->setCellValue($col . '1', '计算机辅助比对审查意见');
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(35);
                    break;
                case 'O':
                    $objSheet->mergeCells('O1:P1');
                    $Spreadsheet->getActiveSheet()->setCellValue($col . '1', '国产替代信息');
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(40);
                    break;
                case 'Q':
                    $objSheet->mergeCells('Q1:V1');
                    $Spreadsheet->getActiveSheet()->setCellValue($col . '1', '用研结合审查专家意见');
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(40);
                    break;
                case 'Y':
                    $objSheet->mergeCells('Y1:Y2');
                    $Spreadsheet->getActiveSheet()->setCellValue($col . '1', '最终审查结论');
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(40);
                    break;
                case 'Z':
                    $objSheet->mergeCells('Z1:Z2');
                    $Spreadsheet->getActiveSheet()->setCellValue($col . '1', '哈希值');
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(40);
                    break;
                default:
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
                    break;
            }

            if ($title == 'kb_substitution_status')
            {
                $objValidation3 = $Spreadsheet->getActiveSheet()->getDataValidation($col . '3');
                $objValidation3->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
                    ->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
                    ->setAllowBlank(false)
                    ->setShowInputMessage(true)
                    ->setShowErrorMessage(true)
                    ->setShowDropDown(true)
                    ->setErrorTitle('输入的值有误')
                    ->setError('您输入的值不在下拉框列表中')
                    ->setFormula1('"成熟产品（CAST/SAST）,成熟产品（字高）,成熟产品（普军/GJB）,成熟产品（COTS）,已鉴定新品,在研新品"');

                $Spreadsheet->getActiveSheet()->setDataValidation("$col" . "4:$col$chunkList", $objValidation3);
            } elseif ($title == 'kb_substitution_whether')
            {
                $objValidation3 = $Spreadsheet->getActiveSheet()->getDataValidation($col . '3');
                $objValidation3->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
                    ->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
                    ->setAllowBlank(false)
                    ->setShowInputMessage(true)
                    ->setShowErrorMessage(true)
                    ->setShowDropDown(true)
                    ->setErrorTitle('输入的值有误')
                    ->setError('您输入的值不在下拉框列表中')
                    ->setFormula1('"原位替代,非原位替代"');
                $Spreadsheet->getActiveSheet()->setDataValidation("$col" . "4:$col$chunkList", $objValidation3);
            } elseif ($title == 'pro_massage')
            {
                $objValidation3 = $Spreadsheet->getActiveSheet()->getDataValidation($col . '3');
                $objValidation3->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
                    ->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
                    ->setAllowBlank(false)
                    ->setShowInputMessage(true)
                    ->setShowErrorMessage(true)
                    ->setShowDropDown(true)
                    ->setErrorTitle('输入的值有误')
                    ->setError('您输入的值不在下拉框列表中')
                    ->setFormula1('"国产替代纳入国产清单,研制攻关暂纳入进口清单,纳入进口清单,不允许选用"');
                $Spreadsheet->getActiveSheet()->setDataValidation("$col" . "4:$col$chunkList", $objValidation3);
            } elseif ($title == 'pro_massage_status')
            {
                $objValidation3 = $Spreadsheet->getActiveSheet()->getDataValidation($col . '3');
                $objValidation3->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
                    ->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
                    ->setAllowBlank(false)
                    ->setShowInputMessage(true)
                    ->setShowErrorMessage(true)
                    ->setShowDropDown(true)
                    ->setErrorTitle('输入的值有误')
                    ->setError('您输入的值不在下拉框列表中')
                    ->setFormula1('"接受,不接受"');
                $Spreadsheet->getActiveSheet()->setDataValidation("$col" . "4:$col$chunkList", $objValidation3);
            } elseif ($title == 'pro_way')
            {
                $objValidation3 = $Spreadsheet->getActiveSheet()->getDataValidation($col . '3');
                $objValidation3->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
                    ->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
                    ->setAllowBlank(false)
                    ->setShowInputMessage(true)
                    ->setShowErrorMessage(true)
                    ->setShowDropDown(true)
                    ->setErrorTitle('输入的值有误')
                    ->setError('您输入的值不在下拉框列表中')
                    ->setFormula1('"采纳国产替代,替换其他国产规格型号,替换其他进口规格型号,不选用,继续选用"');
                $Spreadsheet->getActiveSheet()->setDataValidation("$col" . "4:$col$chunkList", $objValidation3);
            } elseif ($title == 'pro_conclusion')
            {
                $objValidation3 = $Spreadsheet->getActiveSheet()->getDataValidation($col . '3');
                $objValidation3->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
                    ->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
                    ->setAllowBlank(false)
                    ->setShowInputMessage(true)
                    ->setShowErrorMessage(true)
                    ->setShowDropDown(true)
                    ->setErrorTitle('输入的值有误')
                    ->setError('您输入的值不在下拉框列表中')
                    ->setFormula1('"国产替代纳入国产清单,研制攻关暂纳入进口清单,纳入进口清单,替换其他规格型号纳入进口清单,不选用"');
                $Spreadsheet->getActiveSheet()->setDataValidation("$col" . "4:$col$chunkList", $objValidation3);
            }
        }

        // 从0开始的id
        $index = 0;

        // 从第二行开始写数据
        $rows = 2;

        // 方案list
        $cpnModelArr = [
            '认可国产化替代报告方案',
            '补充其他',
            '无替代',
        ];

        foreach ($lists as $k => $list)
        {
            $list  = current($list);
            $index += 1;
            $rows  += 1;

            $cpnModelArr = array_merge($cpnModelArr, array_column($list['auxiliary_result_replace'], 'replace_cpn_specification_model'));

            foreach ($titleKey as $i => $title)
            {
                $cols = $excel->IntToChr($i);

                if ($cols == 'A')
                {
                    $data[$index][$cols . $rows] = $index;
                    $Spreadsheet->getActiveSheet()->setCellValue($cols . $rows, $index);
                } else
                {
                    $Spreadsheet->getActiveSheet()->setCellValue($cols . $rows, @$list[$title]);
                }

                if ($title == 'kb_substitution_plan')
                {
                    if (!empty($cpnModelArr))
                    {
                        $str           = implode(',', $cpnModelArr);
                        $objValidation = $Spreadsheet->getActiveSheet()->getCell($cols . $rows)->getDataValidation();
                        $objValidation->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST);
                        $objValidation->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION);
                        $objValidation->setAllowBlank(false);
                        $objValidation->setShowInputMessage(true);
                        $objValidation->setShowErrorMessage(true);
                        $objValidation->setShowDropDown(true);
                        $objValidation->setErrorTitle('输入的值有误');
                        $objValidation->setError('您输入的值不在下拉框列表中');
                        $objValidation->setFormula1('"' . $str . '"');
                    }
                }
            }
        }

        try
        {
            $objWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($Spreadsheet, 'Xls');
            $filename  = $fileName . '审查表' . date('YmdHis') . '.xls';
            $filename_ = $excelDir . $filename;
            $objWriter->save($filename_);

            return [
                'status'    => true,
                'file_path' => $filename_,
                'file_name' => $filename
            ];
        } catch (\Exception $e)
        {
            return [$e];
        }
    }

    /**
     * 导出审查补充清单02
     * @param $excelDir
     * @param $fileName
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function exportExcelExt($excelDir, $fileName)
    {
        $excel = new Excel();

        $excelTitle01 = $excel->importTitleExt01;

        $titleKey01 = array_keys($excelTitle01);

        // excel处理
        $Spreadsheet = new Spreadsheet();
        $Spreadsheet->setActiveSheetIndex(0);
        $objSheet = $Spreadsheet->getActiveSheet();
        $objSheet->setTitle('国产替代信息补充表');

        // 表头
        $Spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight('30');
        $Spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight('40');

        foreach ($titleKey01 as $i => $title)
        {
            $col = $excel->IntToChr($i);

            $Spreadsheet->getActiveSheet()->setCellValue($col . '2', @$excelTitle01[$title][0]);

            $Spreadsheet->getActiveSheet()->getStyle($col . '2')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '2')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '2')->getAlignment()->setWrapText(true);

            $Spreadsheet->getActiveSheet()->getStyle($col . '1')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '1')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '1')->getAlignment()->setWrapText(true);

            // 设置列宽,默认10
            switch ($col)
            {
                case 'A':
                    $objSheet->mergeCells('A1:B1');
                    $Spreadsheet->getActiveSheet()->setCellValue($col . '1', '被替代进口电子元器件信息');
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(30);
                    break;
                case 'B':
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(30);
                    break;
                case 'C':
                    $objSheet->mergeCells('C1:O1');
                    $Spreadsheet->getActiveSheet()->setCellValue($col . '1', '国产替代补充信息');
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
                    break;
                case 'D':
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
                    break;
                case 'I':
                    $objValidation3 = $Spreadsheet->getActiveSheet()->getDataValidation($col . '3');
                    $objValidation3->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
                        ->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
                        ->setAllowBlank(false)
                        ->setShowInputMessage(true)
                        ->setShowErrorMessage(true)
                        ->setShowDropDown(true)
                        ->setErrorTitle('输入的值有误')
                        ->setError('您输入的值不在下拉框列表中')
                        ->setFormula1('"A,B,C,D,E"');
                    $Spreadsheet->getActiveSheet()->setDataValidation("$col" . "3:" . $col . "100", $objValidation3);
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
                    break;
                case 'M':
                    $objValidation3 = $Spreadsheet->getActiveSheet()->getDataValidation($col . '3');
                    $objValidation3->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
                        ->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
                        ->setAllowBlank(false)
                        ->setShowInputMessage(true)
                        ->setShowErrorMessage(true)
                        ->setShowDropDown(true)
                        ->setErrorTitle('输入的值有误')
                        ->setError('您输入的值不在下拉框列表中')
                        ->setFormula1('"原位替代,非原位替代"');
                    $Spreadsheet->getActiveSheet()->setDataValidation("$col" . "3:" . $col . "100", $objValidation3);
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
                    break;
                case 'N':
                    $objValidation3 = $Spreadsheet->getActiveSheet()->getDataValidation($col . '3');
                    $objValidation3->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
                        ->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
                        ->setAllowBlank(false)
                        ->setShowInputMessage(true)
                        ->setShowErrorMessage(true)
                        ->setShowDropDown(true)
                        ->setErrorTitle('输入的值有误')
                        ->setError('您输入的值不在下拉框列表中')
                        ->setFormula1('"成熟产品（CAST/SAST）,成熟产品（字高）,成熟产品（普军/GJB）,成熟产品（COTS）,已鉴定新品,在研新品"');
                    $Spreadsheet->getActiveSheet()->setDataValidation("$col" . "3:" . $col . "100", $objValidation3);
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
                    break;
                default:
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
                    break;
            }


        }

        $Spreadsheet->createSheet();
        $Spreadsheet->setActiveSheetIndex(1)->setTitle('进口替代信息补充表');
        $objSheet     = $Spreadsheet->getActiveSheet();
        $excelTitle02 = $excel->importTitleExt02;

        $titleKey02 = array_keys($excelTitle02);

        // 表头
        $Spreadsheet->getActiveSheet()->getRowDimension('1')->setRowHeight('30');
        $Spreadsheet->getActiveSheet()->getRowDimension('2')->setRowHeight('40');

        foreach ($titleKey02 as $i => $title)
        {
            $col = $excel->IntToChr($i);
            $Spreadsheet->getActiveSheet()->setCellValue($col . '2', @$excelTitle02[$title][0]);

            $Spreadsheet->getActiveSheet()->getStyle($col . '2')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '2')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '2')->getAlignment()->setWrapText(true);

            $Spreadsheet->getActiveSheet()->getStyle($col . '1')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '1')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $Spreadsheet->getActiveSheet()->getStyle($col . '1')->getAlignment()->setWrapText(true);

            // 设置列宽,默认10
            switch ($col)
            {
                case 'A':
                    $objSheet->mergeCells('A1:B1');
                    $Spreadsheet->getActiveSheet()->setCellValue($col . '1', '被替代进口电子元器件信息');
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(30);
                    break;
                case 'B':
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(30);
                    break;
                case 'C':
                    $objSheet->mergeCells('C1:O1');
                    $Spreadsheet->getActiveSheet()->setCellValue($col . '1', '进口替代补充信息');
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
                    break;
                case 'J':
                    $objValidation3 = $Spreadsheet->getActiveSheet()->getDataValidation($col . '3');
                    $objValidation3->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
                        ->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
                        ->setAllowBlank(false)
                        ->setShowInputMessage(true)
                        ->setShowErrorMessage(true)
                        ->setShowDropDown(true)
                        ->setErrorTitle('输入的值有误')
                        ->setError('您输入的值不在下拉框列表中')
                        ->setFormula1('"红色,紫色,橙色,黄色,绿色"');
                    $Spreadsheet->getActiveSheet()->setDataValidation("$col" . "3:" . $col . "100", $objValidation3);
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
                    break;
                case 'K':
                    $objValidation3 = $Spreadsheet->getActiveSheet()->getDataValidation($col . '3');
                    $objValidation3->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
                        ->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
                        ->setAllowBlank(false)
                        ->setShowInputMessage(true)
                        ->setShowErrorMessage(true)
                        ->setShowDropDown(true)
                        ->setErrorTitle('输入的值有误')
                        ->setError('您输入的值不在下拉框列表中')
                        ->setFormula1('"红色,紫色,橙色,黄色,绿色"');
                    $Spreadsheet->getActiveSheet()->setDataValidation("$col" . "3:" . $col . "100", $objValidation3);
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
                    break;
                case 'L':
                    $objValidation3 = $Spreadsheet->getActiveSheet()->getDataValidation($col . '3');
                    $objValidation3->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
                        ->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
                        ->setAllowBlank(false)
                        ->setShowInputMessage(true)
                        ->setShowErrorMessage(true)
                        ->setShowDropDown(true)
                        ->setErrorTitle('输入的值有误')
                        ->setError('您输入的值不在下拉框列表中')
                        ->setFormula1('"1,2.1,2.2,2.3,3,4"');
                    $Spreadsheet->getActiveSheet()->setDataValidation("$col" . "3:" . $col . "100", $objValidation3);
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
                    break;
                case 'O':
                    $objValidation3 = $Spreadsheet->getActiveSheet()->getDataValidation($col . '3');
                    $objValidation3->setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
                        ->setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
                        ->setAllowBlank(false)
                        ->setShowInputMessage(true)
                        ->setShowErrorMessage(true)
                        ->setShowDropDown(true)
                        ->setErrorTitle('输入的值有误')
                        ->setError('您输入的值不在下拉框列表中')
                        ->setFormula1('"直接,第三方,秘密"');
                    $Spreadsheet->getActiveSheet()->setDataValidation("$col" . "3:" . $col . "100", $objValidation3);
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
                    break;
                default:
                    $Spreadsheet->getActiveSheet()->getColumnDimension($col)->setWidth(20);
                    break;
            }
        }
        try
        {
            $objWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($Spreadsheet, 'Xls');
            $filename  = $fileName . '审查报告' . date('YmdHis') . '.xls';
            $filename_ = $excelDir . $filename;
            $objWriter->save($filename_);

            return [
                'status'    => true,
                'file_path' => $filename_,
                'file_name' => $filename
            ];
        } catch (\Exception $e)
        {
            return [$e];
        }
    }

    /**
     * 导出审查补充清单03
     * @param        $lists
     * @param string $excelDir
     * @param string $fileName
     */
    public function exportWord($lists, $excelDir = '', $fileName = '')
    {
        $phpWord = new PhpWord();

        // 设置默认样式
        $phpWord->setDefaultFontName('宋体');
        $phpWord->setDefaultFontSize(16);

        // 添加页面
        $section = $phpWord->createSection();

        $phpWord->addFontStyle('toc04', ['name' => '宋体', 'size' => 15, 'color' => 'black', 'bold' => true, 'spaceAfter' => 20]);
        $phpWord->addFontStyle('toc01', ['name' => '宋体', 'size' => 26, 'color' => 'black', 'bold' => true, 'spaceAfter' => 20]);
        $phpWord->addFontStyle('toc02', ['name' => '宋体', 'size' => 17, 'color' => 'black', 'bold' => true, 'spaceAfter' => 20]);
        $phpWord->addFontStyle('toc03', ['name' => '宋体', 'size' => 13, 'color' => 'black', 'bold' => false, 'spaceAfter' => 20]);
        $phpWord->addFontStyle('toc031', ['name' => '宋体', 'size' => 10, 'color' => 'black', 'bold' => false, 'spaceAfter' => 20]);

        $styleTable01 = array('borderSize' => 6, 'borderColor' => 'ffffff', 'cellMargin' => 80, 'alignment' => JcTable::CENTER);
        $styleTable02 = array('borderSize' => 6, 'borderColor' => '000000', 'cellMargin' => 80, 'alignment' => JcTable::CENTER);
        $phpWord->addTableStyle('table_01', $styleTable01);
        $phpWord->addTableStyle('table_02', $styleTable02);

        $section->addTextBreak(10);

        $section->addText($fileName . '<w:br/>审查报告', 'toc01', ['align' => 'center']);

        $section->addTextBreak(3);

        $table = $section->addTable('table_01');
        $table->addRow(25);
        $table->addCell(2000, ['borderBottomColor' => 'ffffff'])->addText("委托单位 :", 'toc04', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell(2000)->addText("", 'toc04', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addRow(25);
        $table->addCell(2000, ['borderBottomColor' => 'ffffff'])->addText("审查机构 :", 'toc04', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell(2000)->addText("", 'toc04', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addRow(25);
        $table->addCell(2000, ['borderBottomColor' => 'ffffff'])->addText("主    审 :", 'toc04', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell(2000)->addText("", 'toc04', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addRow(25);
        $table->addCell(2000, ['borderBottomColor' => 'ffffff'])->addText("审    核 :", 'toc04', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell(2000)->addText("", 'toc04', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addRow(25);
        $table->addCell(2000, ['borderBottomColor' => 'ffffff'])->addText("批    准 :", 'toc04', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell(2000)->addText("", 'toc04', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);

        $section->addTextBreak(3);

        $section->addText(date('Y年m月d日'), 'toc02', ['align' => 'center']);

        $section      = $phpWord->createSection();
        $sectionStyle = $section->getSettings();
        $sectionStyle->setLandscape();

        $index = 0;

        foreach ($lists as $key => $list)
        {
            $index++;
            $list = current($list);

            $section->addText($index . '. 规格型号' . $list['cpn_specification_model'] . '、生产厂商' . $list['cpn_manufacturer'], 'toc02');
            $section->addText('    1、计算机辅助比对审查结果：进入用研结合审查', 'toc03');
            $section->addText('    2、最终审查结论：（请补充填写）', 'toc03');
            $section->addText('    3、国产替代补充信息：', 'toc03');
            $section->addText('    （最终审查结论为“国产替代纳入国产清单“时按附表1格式补充填写；否则填“无”）', 'toc03');
            $section->addText('    4、进口替换补充信息：', 'toc03');
            $section->addText('    （最终审查结论为“替换其他规格型号纳入进口清单”时按附表2格式补充填写；否则填“无”）', 'toc03');

            $section->addTextBreak(1);
        }

        $section      = $phpWord->createSection();
        $sectionStyle = $section->getSettings();
        $sectionStyle->setLandscape();

        $section->addTextBreak(2);

        $width0201 = 1100;

        $section->addText('附表1：国产替代信息补充表', 'toc02');
        $table = $section->addTable('table_02');
        $table->addRow(25);
        $table->addCell($width0201)->addText("分类代码", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("电子元器件名称", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("型号规格", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("生产厂商", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("质量等级", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("封装形式", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("自主可控等级", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("参考价格（元）", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("供货周期（周）", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件检测机构", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("替代类型", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("替代产品状态", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addRow(25);
        $table->addCell($width0201)->addText("按GJB8118-2013《军用电子元器件分类与代码》执行", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件供货时的名称，填写元器件中文全称", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件完整型号规格", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件生产厂商名称", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件生产控制质量等级，如GJB597B中的S、BG、B级，GJB2438B的K、H、G、D级等", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件为标准封装的填写封装形式代码，如：CERDIP8、TSOP48、BGA144等；非标准封装填写外形尺寸，如：10×10×2.2，单位默认为mm", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("芯片类产品按照GJB9530-2018《军用关键软硬件自主可控评估通用准则》执行，非芯片类产品参照执行，包括A、B、C、D、E", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件的采购单价，以人民币元为单位填写，可以填写价格范围", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件生产供货整个周期所需的时间，以周为单位填写，可以填写供货周期范围", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件鉴定检测和保障机构名称", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("包括：原位替代、非原位替代", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("包括：成熟产品(CAST/SAST)、成熟产品(宇高)、成熟产品(普军/GBJ)、成熟产品(COTS)、已鉴定新品、在研新品", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);

        $section      = $phpWord->createSection();
        $sectionStyle = $section->getSettings();
        $sectionStyle->setLandscape();

        $section->addTextBreak(2);

        $section->addText('附表2：进口替换信息补充表', 'toc02');
        $table = $section->addTable('table_02');
        $table->addRow(25);
        $table->addCell($width0201)->addText("分类代码", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("电子元器件名称", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("型号规格", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("生产厂商", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("国别地区", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("质量等级", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("封装形式", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("安全等级颜色", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("建议安全等级颜色", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("必要性", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("参考价格（元）", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("供货周期（周）", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("获取渠道", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addRow(25);
        $table->addCell($width0201)->addText("按GJB8118-2013《军用电子元器件分类与代码》执行", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件供货时的名称，填写元器件中文全称", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件完整型号规格", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件生产厂商名称，勿填写国内代理商", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件生产厂商所属国别地区，如美国、英国、中国台湾等", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件生产控制质量等级，如MIL-PRF-38535 Q级、MIL-PRF-38534 H级、工业级、商业级等", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件为标准封装的填写封装形式代码，如：CERDIP8、TSOP48、BGA144等；非标准封装填写外形尺寸，如：10×10×2.2，单位默认为mm", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("包括：红色、紫色、橙色、黄色和绿色", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("包括：红色、紫色、橙色、黄色和绿色", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("没有类似国产元器件填写“1”；国产类似元器件性能指标达不到使用要求填写“2.1”；国产类似元器件可靠性指标达不到使用要求填写“2.2”；国产类似元器件体积/重量达不到使用要求填写“2.3”；国产类似元器件价格昂贵填写“3”", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件的采购单价，以人民币元为单位填写，可以填写价格范围", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("元器件生产供货整个周期所需的时间，以周为单位填写，可以填写供货周期范围", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);
        $table->addCell($width0201)->addText("直接进口贸易渠道的填“直接”；通过第三方转口贸易规避禁运隐蔽获取的填“第三方”；秘密渠道的填写“秘密”；为选填项", 'toc031', ['align' => 'left', 'valign' => 'center', 'color' => '4d79ff']);

        try
        {
            // 如果文件夹不存在 则创建
            if (!file_exists(public_path('attachment/outword/')))
            {
                mkdir(public_path('attachment/outword/'), '0777', true);
            }

            // 导出文件
            $filename  = $fileName . '审查报告' . date('YmdHis') . '.docx';
            $filename_ = $excelDir . $filename;
            $writer    = IOFactory::createWriter($phpWord, 'Word2007');

            $writer->save($filename_);
            return [
                'status'    => true,
                'file_path' => $filename_,
                'file_name' => $filename
            ];
        } catch (\Exception $e)
        {
            return [$e];
        }
    }

    /**
     * 生成错误信息的excel
     * @return array
     * @throws \PHPExcel_Exception
     */
    public function exportError($errData, $excelDir, $fileName)
    {
        $excel = new Excel();

        $excelTitle01 = [
            'id'          => ['序号'],
            'error_local' => ['错误位置'],
            'error_msg'   => ['错误描述'],
        ];

        $titleKey = array_keys($excelTitle01);

        // excel处理
        $objExcel = new \PHPExcel();

        $objExcel->setActiveSheetIndex(0);
        $objSheet = $objExcel->getActiveSheet();
        $objSheet->setTitle('错误信息表');

        // 表头
        $objExcel->getActiveSheet()->getRowDimension('1')->setRowHeight('30');

        foreach ($titleKey as $i => $title)
        {
            $col = $excel->IntToChr($i);
            $objExcel->getActiveSheet()->setCellValue($col . '1', @$excelTitle01[$title][0]);
            $objExcel->getActiveSheet()->getStyle($col . '1:' . $col . '100')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
            $objExcel->getActiveSheet()->getStyle($col . '1:' . $col . '100')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $objExcel->getActiveSheet()->getStyle($col . '1:' . $col . '100')->getAlignment()->setWrapText(true);
            switch ($col)
            {
                case 'A':
                    $objExcel->getActiveSheet()->getColumnDimension($col)->setWidth(50);
                    break;
                case 'B':
                    $objExcel->getActiveSheet()->getColumnDimension($col)->setWidth(50);
                    break;
                case 'C':
                    $objExcel->getActiveSheet()->getColumnDimension($col)->setWidth(50);
                    break;
                default:

                    break;
            }
        }

        $index = 0;
        $rows  = 1;

        foreach ($errData as $k => $list)
        {
            $index += 1;
            $rows  += 1;

            foreach ($titleKey as $i => $title)
            {
                $cols = $excel->IntToChr($i);

                if ($cols == 'A')
                {
                    $objExcel->getActiveSheet()->setCellValue($cols . $rows, $index);
                } else
                {
                    $objExcel->getActiveSheet()->setCellValue($cols . $rows, $list[$title]);
                }
            }
        }


        try
        {
            $objWriter = \PHPExcel_IOFactory::createWriter($objExcel, 'Excel5');
            $filename  = $fileName . '审查报告错误信息' . date('YmdHis') . '.xls';
            $filename_ = $excelDir . $filename;
            $objWriter->save($filename_);

            return [
                'status'    => true,
                'file_path' => $filename_,
                'file_name' => $filename
            ];
        } catch (\Exception $e)
        {
            return [$e];
        }
    }

    /**
     * 解压文件夹
     * @param $rootPath
     * @param $file
     * @return array
     */
    public function unzip($rootPath, $file)
    {
        try
        {
            $file_name = $file->getClientOriginalName();
            $file->move($rootPath, $file_name);

            $workPath  = $rootPath . DIRECTORY_SEPARATOR . str_replace('.zip', '', $file_name);
            $file_path = $rootPath . DIRECTORY_SEPARATOR . $file_name;

            $modelStruct = new ModelStructure();
            $modelStruct->unzip($file_path, $workPath);

            $fileList = [];
            $dir      = opendir($workPath . DIRECTORY_SEPARATOR . 'package');

            while ($file = readdir($dir))
            {
                array_push($fileList, $file);
            }

            closedir();

            array_shift($fileList);
            array_shift($fileList);

            return [
                'status' => true,
                'data'   => $fileList
            ];

        } catch (\Exception $exception)
        {
            return [
                'status' => false,
                'data'   => $exception
            ];
        }

    }

    /**
     * 验证文件完整性 可用性
     * @param $unzipResult
     * @param $fileName
     * @return array
     */
    public function verifyFile($unzipResult, $fileName)
    {
        // 判断是不是三个文件
        if (count($unzipResult) !== 3)
            return [
                'status' => false,
                'msg'    => '包中文件数量不正确'
            ];

        // 验证是不是两个excel和一个word
        $verifyExt = [
            0 => 'xls',
            1 => 'xls',
            2 => 'docx',
        ];

        foreach ($unzipResult as $key => $value)
        {
            $arr = explode('.', $value);
            if (count($arr) != 2)
                return [
                    'status' => false,
                    'msg'    => '文件名称不能修改'
                ];


            $ext = end($arr);
            if (in_array($ext, $verifyExt))
            {
                $key = array_key_exists($ext, $verifyExt);
                unset($verifyExt[$key]);
            }
        }

        if (empty($verifyExt))
            return [
                'status' => false,
                'msg'    => '包中缺少文件或者文件类型不正确'
            ];

        // 判断文件名称上的型号名称，研制阶段是否正确
        foreach ($unzipResult as $key => $value)
        {
            $arr = explode('.', $value);
            if (count($arr) != 2)
                return [
                    'status' => false,
                    'msg'    => '文件名称不能修改'
                ];


            $name     = current($arr);
            $matchRes = preg_match("/^" . $fileName . "/", $name);
            if (!$matchRes)
                return [
                    'status' => false,
                    'msg'    => '文件名称中的型号和研制阶段匹配失败'
                ];
        }

        return [
            'status' => true,
            'msg'    => '验证成功'
        ];
    }

    /**
     * 获取想要的文件名称
     * @param $fileList
     * @param $fileName
     * @return array
     */
    public function getFileName($fileList, $fileName, $matchName)
    {
        foreach ($fileList as $key => $value)
        {
            $arr = explode('.', $value);
            if (count($arr) != 2)
                return [
                    'status' => false,
                    'msg'    => '文件名称不能修改'
                ];


            $name     = current($arr);
            $ext      = end($arr);
            $matchRes = preg_match("/^" . $fileName . $matchName . "/", $name);
            if ($matchRes == true && $ext == 'xls')
                return [
                    'status' => true,
                    'msg'    => $value
                ];
        }

        return [
            'status' => false,
            'msg'    => '未找到对应文件'
        ];
    }

    /**
     * 验证文件是否可读
     * @param $readExcelResult
     * @return array
     */
    public function verifyExcelData01($readExcelResult)
    {

        // 查看是不是一个sheet
        if (count($readExcelResult['data']) != 1)
            return [
                'status' => false,
                'msg'    => '不是一个sheet'
            ];

        // 查看表头是不是26列
        $data = current($readExcelResult['data']);
        if (count($data['header']) != 26)
            return [
                'status' => false,
                'msg'    => '表头不正确不是26列'
            ];

        // 判断表头是否被处理过
        $titleArr = [
            "A" => "",
            "B" => "",
            "C" => "元器件类别",
            "D" => "元器件名称",
            "E" => "规格型号",
            "F" => "生产厂商",
            "G" => "国别地区",
            "H" => "质量等级",
            "I" => "封装形式",
            "J" => "安全颜色等级",
            "K" => "建议安全等级颜色",
            "L" => "是否核心关键器件",
            "M" => "必要性",
            "N" => "",
            "O" => "可替代产品型号(厂家/质量等级)",
            "P" => "可替代产品应用状态",
            "Q" => "替代方案选择(选择一种最优方案)",
            "R" => "可替代产品型号(选择认可国产化替代报告方案或补充其他时填写)",
            "S" => "可替代产品厂商(选择认可国产化替代报告方案或补充其他时填写)",
            "T" => "替代类型",
            "U" => "可替代产品状态",
            "V" => "专家审查结果",
            "W" => "是否接受专家(或计算机比对)审查意见",
            "X" => "具体处理措施",
            "Y" => "",
            "Z" => "",
        ];

        if (!empty(array_diff(current($data['rowsData']), $titleArr)))
            return [
                'status' => false,
                'msg'    => '表头不正确'
            ];

        return [
            'status' => true,
            'msg'    => '验证成功'
        ];
    }

    /**
     * 验证数据是否可读
     * @param $readExcelResult
     * @return array
     */
    public function verifyExcelData02($readExcelResult)
    {
        // 查看是不是一个sheet
        if (count($readExcelResult['data']) != 2)
            return [
                'status' => false,
                'msg'    => '不是两个sheet'
            ];

        // 查看表头是不是26列
        $data01 = current($readExcelResult['data']);
        if (count($data01['header']) != 14)
            return [
                'status' => false,
                'msg'    => 'SHEET1表头不正确不是15列'
            ];

        $data02 = end($readExcelResult['data']);
        if (count($data02['header']) != 15)
            return [
                'status' => false,
                'msg'    => 'SHEET2表头不正确不是16列'
            ];

        // 判断表头是否被处理过
        $titleArr01 = [
            "A" => "用研审查清单序号",
            "B" => "型号规格",
            "C" => "分类代码",
            "D" => "元器件名称",
            "E" => "替代规格型号",
            "F" => "替代生产厂商",
            "G" => "质量等级",
            "H" => "封装形式",
            "I" => "自主可控等级",
            "J" => "参考价格",
            "K" => "供货周期",
            "L" => "元器件检测机构",
            "M" => "替代类型",
            "N" => "产品状态",
        ];

        $titleArr02 = [
            "A" => "用研审查清单序号",
            "B" => "型号规格",
            "C" => "分类代码",
            "D" => "元器件名称",
            "E" => "替代规格型号",
            "F" => "替代生产厂商",
            "G" => "国别地区",
            "H" => "质量等级",
            "I" => "封装形式",
            "J" => "安全等级颜色",
            "K" => "建议安全等级颜色",
            "L" => "必要性",
            "M" => "参考价格",
            "N" => "供货周期",
            "O" => "获取渠道",
        ];

        if (!empty(array_diff(current($data01['rowsData']), $titleArr01)))
            return [
                'status' => false,
                'msg'    => 'SHEET1表头不正确'
            ];

        if (!empty(array_diff(current($data02['rowsData']), $titleArr02)))
            return [
                'status' => false,
                'msg'    => 'SHEET2表头不正确'
            ];

        return [
            'status' => true,
            'msg'    => '验证成功'
        ];
    }

    /**
     * 处理导出数据
     * @param $readExcelResult
     * @return array
     */
    public function exportExcelData($main_task_id, $readExcelResult01, $readExcelResult02)
    {
        $okArr = [
            '国产替代纳入国产清单',
            '替换其他规格型号纳入进口清单',
        ];

        $cpnImport = new CpnImport();

        $cpnModelArr = [];
        $updateArr   = [];

        // 获取第一个excel数据 删除表头数据
        $data01 = current($readExcelResult01['data'])['rowsData'];
        array_shift($data01);

        // 循环数据
        foreach ($data01 as $key => $row)
        {
            // 更新条件
            $update['hash_code'] = $row['Z'];

            // 更新字段
            $update['kb_substitution_plan']    = $row['Q'];
            $update['kb_substitution_model']   = $row['R'];
            $update['kb_substitution_mfr']     = $row['S'];
            $update['kb_substitution_whether'] = $row['T'];
            $update['kb_substitution_status']  = $row['U'];
            $update['pro_massage']             = $row['V'];
            $update['pro_massage_status']      = $row['W'];
            $update['pro_way']                 = $row['X'];
            $update['pro_conclusion']          = $row['Y'];
            array_push($updateArr, $update);

            // 获取最终结论为$okArr中的数值时 保存该元器件的型号名称
            if (in_array($row['Y'], $okArr))
                array_push($cpnModelArr, $row['D']);
        }

        // 补充excel sheet1 删除表头数据
        $sheet1 = current($readExcelResult02['data'])['rowsData'];
        array_shift($sheet1);

        // 补充excel sheet2 删除表头数据
        $sheet2 = end($readExcelResult02['data'])['rowsData'];
        array_shift($sheet2);

        // 国产意见补充表的数据
        $sheetList = [];

        foreach ($sheet1 as $key => $value)
        {
            if (in_array($value['B'], $cpnModelArr))
                $sheet1Data['is_pass'] = 1;
            else
                $sheet2Data['is_pass'] = 0;

            $sheet1Data['hash_code']                       = $value['A'];
            $sheet1Data['type']                            = 1;
            $sheet1Data['cpn_specification_model']         = $value['B'];
            $sheet1Data['cpn_manufacturer']                = $value['C'];
            $sheet1Data['cpn_manufacturer_replace']        = $value['F'];
            $sheet1Data['cpn_specification_model_replace'] = $value['G'];
            $sheet1Data['cpn_category_code']               = $value['D'];
            $sheet1Data['cpn_name']                        = $value['E'];
            $sheet1Data['cpn_quality']                     = $value['H'];
            $sheet1Data['cpn_package']                     = $value['I'];
            $sheet1Data['cpn_control_level']               = $value['J'];
            $sheet1Data['cpn_ref_price']                   = $value['K'];
            $sheet1Data['cpn_period']                      = $value['L'];
            $sheet1Data['cpn_detect_apartment']            = $value['M'];
            $sheet1Data['cpn_status']                      = $value['O'];
            $sheet1Data['cpn_replace_status']              = $value['N'];

            array_push($sheetList, $sheet1Data);
        }

        foreach ($sheet2 as $key => $value)
        {
            if (in_array($value['B'], $cpnModelArr))
                $sheet2Data['is_pass'] = 1;
            else
                $sheet2Data['is_pass'] = 0;

            $sheet2Data['hash_code']                       = $value['A'];
            $sheet2Data['type']                            = 2;
            $sheet2Data['cpn_specification_model']         = $value['B'];
            $sheet2Data['cpn_manufacturer']                = $value['C'];
            $sheet2Data['cpn_manufacturer_replace']        = $value['F'];
            $sheet2Data['cpn_specification_model_replace'] = $value['G'];
            $sheet2Data['cpn_category_code']               = $value['D'];
            $sheet2Data['cpn_name']                        = $value['E'];
            $sheet2Data['cpn_country']                     = $value['H'];
            $sheet2Data['cpn_quality']                     = $value['I'];
            $sheet2Data['cpn_package']                     = $value['J'];
            $sheet2Data['safe_color']                      = $value['K'];
            $sheet2Data['proposed_safe_color']             = $value['L'];
            $sheet2Data['necessity']                       = $value['M'];
            $sheet2Data['cpn_ref_price']                   = $value['N'];
            $sheet2Data['cpn_period']                      = $value['O'];
            $sheet2Data['access_channel']                  = $value['P'];

            array_push($sheetList, $sheet2Data);
        }

        // 检查文件中的数据是否有不正确的地方
        $errData = $this->VerifyExcelDataPass($main_task_id, $updateArr, $sheetList, $cpnModelArr);

        // 当发现错误的时候返回错误 并且将统计到的错误统计到表中
        if (!empty($errData))
        {
            DB::beginTransaction();
            try
            {
                ProReviewError::where('main_task_id', $main_task_id)->delete();
                ProReviewError::insert($errData);
                MainTask::where('id', $main_task_id)->update(['professor_status_ext' => 2]);

                DB::commit();
                return [
                    'status' => true,
                    'msg'    => '上传失败，文件内容有误！具体请下载文件查看！'
                ];
            } catch (\Exception $exception)
            {
                DB::rollBack();
                return [
                    'status' => false,
                    'msg'    => '错误信息更新失败'
                ];
            }
        }

        // 执行sql
        DB::beginTransaction();
        try
        {
            // 更新意见表数据
            ProReviewOpinionAdd::insert($sheetList);
            MainTask::where('id', $main_task_id)->update(['professor_status_ext' => 3]);

            $updateRes01 = $cpnImport->updateBatch('pro_review_opinion_import', $updateArr);
            if ($updateRes01 === false)
            {
                DB::rollBack();
                return [
                    'status' => false,
                    'msg'    => '更新pro_review_opinion_import表失败'
                ];
            }

            $cpnList     = ProReviewOpinionImport::where('main_task_id', $main_task_id)->select('cpn_id', 'pro_conclusion')->get()->toArray();
            $updateRes02 = $cpnImport->updateBatch('cpn_import', $cpnList);
            if ($updateRes02 === false)
            {
                DB::rollBack();
                return [
                    'status' => false,
                    'msg'    => '更新cpn_import表失败'
                ];
            }

            DB::commit();
            return [
                'status' => true,
                'msg'    => '更新成功'
            ];
        } catch (\Exception $exception)
        {
            DB::rollBack();
            return [
                'status' => false,
                'msg'    => $exception
            ];
        }
    }

    /**
     * 组装错误信息
     * @param $data1
     * @param $data2
     * @return array
     */
    public function VerifyExcelDataPass($main_task_id, $data1, $data2, $cpnModelArr)
    {
        $errData = [];
        $listArr = [
            '国产替代纳入国产清单',
            '研制攻关暂纳入进口清单',
            '纳入进口清单',
            '替换其他规格型号纳入进口清单',
            '不选用',
        ];

        $cpn_replace_status_list = [
            '原位替代',
            '非原位替代'
        ];

        $cpn_status_list = [
            '成熟产品（CAST/SAST）',
            '成熟产品（字高）',
            '成熟产品（普军/GJB）',
            '成熟产品（COTS）',
            '已鉴定新品',
            '在研新品',
        ];

        $cpn_control_level_list = [
            'A',
            'B',
            'C',
            'D',
            'E',
        ];

        $safe_color_list = [
            '红色',
            '紫色',
            '橙色',
            '黄色',
            '绿色',
        ];

        $proposed_safe_color_list = [
            '红色',
            '紫色',
            '橙色',
            '黄色',
            '绿色',
        ];

        $necessity_list = [
            '1',
            '2.1',
            '2.2',
            '2.3',
            '3',
            '4',
        ];

        $access_channel_list = [
            '直接',
            '第三方',
            '秘密',
        ];

        $catList = CpnCategories::pluck('id')->toArray();

        // 审查意见表X行X列，错误描述：最终审查结论为空/最终审查结论内容不合规
        foreach ($data1 as $key1 => $value1)
        {
            if (empty($value1['pro_conclusion']))
            {
                $err = [
                    'main_task_id' => $main_task_id,
                    'error_local'  => ($key1 + 4) . '行Y列',
                    'error_msg'    => '最终审查结论为空'
                ];
                array_push($errData, $err);
            } else
            {
                if (!in_array($value1['pro_conclusion'], $listArr))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => ($key1 + 4) . '行Y列',
                        'error_msg'    => '最终审查结论内容不合规'
                    ];
                    array_push($errData, $err);
                }
            }


        }

        // 审查报告1.X规格型号XX、生产厂商XX，错误描述：缺少国产替代补充信息/进口替换补充信息
        $data2CpnModel = [];
        if (!empty($data2))
        {
            $data2CpnModel = array_column('cpn_specification_model', $data2);
        }

        $diffArr = array_diff($cpnModelArr, $data2CpnModel);
        foreach ($diffArr as $key => $value)
        {
            $err = [
                'main_task_id' => $main_task_id,
                'error_local'  => $value . '型号',
                'error_msg'    => '缺少国产替代补充信息/进口替换补充信息'
            ];
            array_push($errData, $err);
        }

        // 审查报告1.X规格型号XX、生产厂商XX，错误描述：国产替代补充信息/进口替换补充信息存在型号规格、生产厂商、分类代码、（替代类型、替代产品状态）缺失，或分类代码不属于GJB 8118，（或替代类型、替代产品状态内容不合规）的问题
        // 审查报告1.X规格型号XX、生产厂商XX，错误描述：国产替代补充信息/进口替换补充信息存在必填信息缺失、填报数据类型不正确或填报信息超出数据字典范围的问题
        foreach ($data2 as $key2 => $value2)
        {
            if ($value2['type'] == 1)
                $sheetName = '国产替代补充信息';
            else
                $sheetName = '进口替换补充信息';

            if (empty($value2['cpn_manufacturer_replace']))
            {
                $err = [
                    'main_task_id' => $main_task_id,
                    'error_local'  => $sheetName($key2 + 4) . '行E列',
                    'error_msg'    => '替代型号规格缺失'
                ];
                array_push($errData, $err);
            }

            if (empty($value2['cpn_name']))
            {
                $err = [
                    'main_task_id' => $main_task_id,
                    'error_local'  => $sheetName($key2 + 4) . '行D列',
                    'error_msg'    => '替代元器件名称缺失'
                ];
                array_push($errData, $err);
            }

            if (empty($value2['cpn_specification_model_replace']))
            {
                $err = [
                    'main_task_id' => $main_task_id,
                    'error_local'  => $sheetName($key2 + 4) . '行F列',
                    'error_msg'    => '替代生产厂商缺失'
                ];
                array_push($errData, $err);
            }

            if (empty($value2['cpn_category_code']))
            {
                $err = [
                    'main_task_id' => $main_task_id,
                    'error_local'  => $sheetName($key2 + 4) . '行C列',
                    'error_msg'    => '分类代码缺失'
                ];
                array_push($errData, $err);
            } else
            {
                if (in_array($value2['cpn_category_code'], $catList))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行C列',
                        'error_msg'    => '分类代码不属于GJB 8118'
                    ];
                    array_push($errData, $err);
                }
            }

            if ($value2['type'] == 1)
            {
                if (empty($value2['cpn_quality']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行G列',
                        'error_msg'    => '质量等级缺失'
                    ];
                    array_push($errData, $err);
                }

                if (empty($value2['cpn_package']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行H列',
                        'error_msg'    => '封装形式缺失'
                    ];
                    array_push($errData, $err);
                }

                if (empty($value2['cpn_replace_status']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行M列',
                        'error_msg'    => '替代类型缺失'
                    ];
                    array_push($errData, $err);
                } else
                {
                    if (in_array($value2['cpn_replace_status'], $cpn_replace_status_list))
                    {
                        $err = [
                            'main_task_id' => $main_task_id,
                            'error_local'  => $sheetName($key2 + 4) . '行M列',
                            'error_msg'    => '替代类型内容不合规'
                        ];
                        array_push($errData, $err);
                    }
                }

                if (empty($value2['cpn_status']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行N列',
                        'error_msg'    => '替代产品状态缺失'
                    ];
                    array_push($errData, $err);
                } else
                {
                    if (in_array($value2['cpn_status'], $cpn_status_list))
                    {
                        $err = [
                            'main_task_id' => $main_task_id,
                            'error_local'  => $sheetName($key2 + 4) . '行N列',
                            'error_msg'    => '替代产品状态内容不合规'
                        ];
                        array_push($errData, $err);
                    }
                }

                if (empty($value2['cpn_control_level']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行I列',
                        'error_msg'    => '自主可控等级缺失'
                    ];
                    array_push($errData, $err);
                } else
                {
                    if (in_array($value2['cpn_control_level'], $cpn_control_level_list))
                    {
                        $err = [
                            'main_task_id' => $main_task_id,
                            'error_local'  => $sheetName($key2 + 4) . '行I列',
                            'error_msg'    => '自主可控等级内容不合规'
                        ];
                        array_push($errData, $err);
                    }
                }

                if (empty($value2['cpn_ref_price']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行J列',
                        'error_msg'    => '参考价格缺失'
                    ];
                    array_push($errData, $err);
                } else
                {
                    if (preg_match('/^[0-9]*$/', $value2['cpn_ref_price']))
                    {
                        $priceArr = explode('_', $value2['cpn_ref_price']);
                        if (count($priceArr) === 2)
                        {
                            if (!preg_match('/^[0-9]*$/', current($priceArr)) || !preg_match('/^[0-9]*$/', end($priceArr)))
                            {
                                $err = [
                                    'main_task_id' => $main_task_id,
                                    'error_local'  => $sheetName($key2 + 4) . '行J列',
                                    'error_msg'    => '参考价格数据类型不正确'
                                ];
                                array_push($errData, $err);
                            }
                        } else
                        {
                            $err = [
                                'main_task_id' => $main_task_id,
                                'error_local'  => $sheetName($key2 + 4) . '行J列',
                                'error_msg'    => '参考价格数据类型不正确'
                            ];
                            array_push($errData, $err);
                        }
                    }
                }

                if (empty($value2['cpn_period']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行K列',
                        'error_msg'    => '供货周期缺失'
                    ];
                    array_push($errData, $err);
                } else
                {
                    if (preg_match('/^[0-9]*$/', $value2['cpn_period']))
                    {
                        $priceArr = explode('_', $value2['cpn_period']);
                        if (count($priceArr) === 2)
                        {
                            if (!preg_match('/^[0-9]*$/', current($priceArr)) || !preg_match('/^[0-9]*$/', end($priceArr)))
                            {
                                $err = [
                                    'main_task_id' => $main_task_id,
                                    'error_local'  => $sheetName($key2 + 4) . '行K列',
                                    'error_msg'    => '供货周期数据类型不正确'
                                ];
                                array_push($errData, $err);
                            }
                        } else
                        {
                            $err = [
                                'main_task_id' => $main_task_id,
                                'error_local'  => $sheetName($key2 + 4) . '行K列',
                                'error_msg'    => '供货周期数据类型不正确'
                            ];
                            array_push($errData, $err);
                        }
                    }
                }


            } else
            {
                if (empty($value2['cpn_quality']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行H列',
                        'error_msg'    => '质量等级缺失'
                    ];
                    array_push($errData, $err);
                }

                if (empty($value2['cpn_package']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行I列',
                        'error_msg'    => '封装形式缺失'
                    ];
                    array_push($errData, $err);
                }

                if (empty($value2['safe_color']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行J列',
                        'error_msg'    => '安全等级颜色缺失'
                    ];
                    array_push($errData, $err);
                } else
                {
                    if (in_array($value2['safe_color'], $safe_color_list))
                    {
                        $err = [
                            'main_task_id' => $main_task_id,
                            'error_local'  => $sheetName($key2 + 4) . '行J列',
                            'error_msg'    => '安全等级颜色内容不合规'
                        ];
                        array_push($errData, $err);
                    }
                }

                if (empty($value2['proposed_safe_color']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行K列',
                        'error_msg'    => '建议安全等级颜色缺失'
                    ];
                    array_push($errData, $err);
                } else
                {
                    if (in_array($value2['proposed_safe_color'], $proposed_safe_color_list))
                    {
                        $err = [
                            'main_task_id' => $main_task_id,
                            'error_local'  => $sheetName($key2 + 4) . '行K列',
                            'error_msg'    => '建议安全等级颜色内容不合规'
                        ];
                        array_push($errData, $err);
                    }
                }

                if (empty($value2['necessity']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行L列',
                        'error_msg'    => '必要性缺失'
                    ];
                    array_push($errData, $err);
                } else
                {
                    if (in_array($value2['necessity'], $necessity_list))
                    {
                        $err = [
                            'main_task_id' => $main_task_id,
                            'error_local'  => $sheetName($key2 + 4) . '行L列',
                            'error_msg'    => '必要性内容不合规'
                        ];
                        array_push($errData, $err);
                    }
                }

                if (empty($value2['access_channel']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行O列',
                        'error_msg'    => '获取渠道缺失'
                    ];
                    array_push($errData, $err);
                } else
                {
                    if (in_array($value2['access_channel'], $access_channel_list))
                    {
                        $err = [
                            'main_task_id' => $main_task_id,
                            'error_local'  => $sheetName($key2 + 4) . '行O列',
                            'error_msg'    => '获取渠道不合规'
                        ];
                        array_push($errData, $err);
                    }
                }

                if (empty($value2['cpn_ref_price']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行M列',
                        'error_msg'    => '参考价格缺失'
                    ];
                    array_push($errData, $err);
                } else
                {
                    if (preg_match('/^[0-9]*$/', $value2['cpn_ref_price']))
                    {
                        $priceArr = explode('_', $value2['cpn_ref_price']);
                        if (count($priceArr) === 2)
                        {
                            if (!preg_match('/^[0-9]*$/', current($priceArr)) || !preg_match('/^[0-9]*$/', end($priceArr)))
                            {
                                $err = [
                                    'main_task_id' => $main_task_id,
                                    'error_local'  => $sheetName($key2 + 4) . '行M列',
                                    'error_msg'    => '参考价格数据类型不正确'
                                ];
                                array_push($errData, $err);
                            }
                        } else
                        {
                            $err = [
                                'main_task_id' => $main_task_id,
                                'error_local'  => $sheetName($key2 + 4) . '行M列',
                                'error_msg'    => '参考价格数据类型不正确'
                            ];
                            array_push($errData, $err);
                        }
                    }
                }

                if (empty($value2['cpn_period']))
                {
                    $err = [
                        'main_task_id' => $main_task_id,
                        'error_local'  => $sheetName($key2 + 4) . '行N列',
                        'error_msg'    => '供货周期缺失'
                    ];
                    array_push($errData, $err);
                } else
                {
                    if (preg_match('/^[0-9]*$/', $value2['cpn_period']))
                    {
                        $priceArr = explode('_', $value2['cpn_period']);
                        if (count($priceArr) === 2)
                        {
                            if (!preg_match('/^[0-9]*$/', current($priceArr)) || !preg_match('/^[0-9]*$/', end($priceArr)))
                            {
                                $err = [
                                    'main_task_id' => $main_task_id,
                                    'error_local'  => $sheetName($key2 + 4) . '行N列',
                                    'error_msg'    => '供货周期数据类型不正确'
                                ];
                                array_push($errData, $err);
                            }
                        } else
                        {
                            $err = [
                                'main_task_id' => $main_task_id,
                                'error_local'  => $sheetName($key2 + 4) . '行N列',
                                'error_msg'    => '供货周期数据类型不正确'
                            ];
                            array_push($errData, $err);
                        }
                    }
                }
            }
        }

        return $errData;
    }

    /**
     * 将数据放到知识库中
     */
    public function insertCissData()
    {
        //        // ciss库要进行修改的数据
        //        $reUpdateList = [];
        //        // ciss库需要新增的数据
        //        $insertList = [];
        //        // ciss_history要进行新增的数据
        //        $insertListExt = [];
        //
        //        $reColumn = [];
        //
        //        $cissData = AuxiliaryDataDomesticCiss::select('id', 'cpn_specification_model', 'replace_cpn_specification_model')
        //            ->get()
        //            ->groupBy('cpn_specification_model')
        //            ->toArray();
        //
        //        if (in_array($value['B'], $cpnModelArr))
        //        {
        //            // 判断是否有这个型号的数据
        //            if (key_exists($value['B'], $cissData))
        //            {
        //                $reData = $cissData[$value['B']];
        //
        //                // 将这个型号的全部替代型号取出
        //                foreach ($reData as $reKey => $reValue)
        //                {
        //                    $reColumn[$reValue['id'] . '|' . $reValue['hash']] = $reValue['replace_cpn_specification_model'];
        //                }
        //
        //                // 判断当前的是否是替代型号
        //                if (in_array($value['F'], $reColumn))
        //                {
        //                    $keyIndex    = array_search($value['F'], $reColumn);
        //                    $keyIndexArr = explode("|", $keyIndex);
        //
        //                    // 组装要修改的数据
        //                    $reData['id']                    = current($keyIndexArr);
        //                    $reData['replace_type']          = $value['N'];
        //                    $reData['replace_product_state'] = $value['O'];
        //                    $reData['type']                  = 1;
        //
        //                    array_push($reUpdateList, $reData);
        //
        //                    // 组装要新增的数据
        //                    $insData['hash']       = end($keyIndexArr);
        //                    $insData['equip_type'] = '';            // todo 需要查询
        //                    $insData['model_name'] = $value['B'];
        //                    $insData['equip_num']  = 0;             // todo 需要计算
        //                    $insData['price']      = $value['K'];
        //                    $insData['delivery']   = $value['L'];
        //                }
        //            } else
        //            {
        //                $hash = sha1(time() . rand(999, 999999) . rand(999, 999999));
        //
        //                // 组装知识库中要新增的数据
        //                $insertData['cpn_name']                        = $value['E'];
        //                $insertData['cpn_specification_model']         = $value['B'];
        //                $insertData['cpn_manufacturer']                = $value['C'];
        //                $insertData['replace_cpn_specification_model'] = $value['F'];
        //                $insertData['replace_cpn_manufacturer']        = $value['G'];
        //                $insertData['replace_cpn_quality']             = $value['H'];
        //                $insertData['replace_type']                    = $value['N'];
        //                $insertData['replace_product_state']           = $value['O'];
        //                $insertData['type']                            = 1;
        //                $insertData['hash']                            = $hash;
        //
        //                array_push($insertList, $insertData);
        //
        //                // 组装知识库补充表的数据
        //                $insData['hash']       = $hash;
        //                $insData['equip_type'] = '';            // todo 需要查询
        //                $insData['model_name'] = $value['B'];
        //                $insData['equip_num']  = 0;             // todo 需要计算
        //                $insData['price']      = $value['K'];
        //                $insData['delivery']   = $value['L'];
        //            }
        //
        //            // 将ciss_history中需要新增的数据全部插入到这个数组中一次性插入
        //            array_push($insertListExt, $insData);
        //        }
        //
        //        // 更新知识库数据
        //        AuxiliaryDataDomesticCissHistory::insert($insertListExt);
        //        AuxiliaryDataDomesticCiss::insert($insertList);
        //
        //        $updateRes02 = $cpnImport->updateBatch('auxiliary_data_domestic_ciss', $reUpdateList);
        //
        //        if ($updateRes02 === false)
        //        {
        //            DB::rollBack();
        //            return [
        //                'status' => false,
        //                'msg'    => '更新auxiliary_data_domestic_ciss表失败'
        //            ];
        //        }
    }

    /**
     * 获取审查后双清单
     * @param $cpn_type
     * @param $main_task_id
     * @return mixed
     */
    public function getCpnData($main_task_id, $cpn_type)
    {
        $where = '';

        $select2 = [
            'cpn_category_code',
            'cpn_manufacturer',
            'cpn_specification_model',
            'cpn_name',
            'cpn_quality',
            'cpn_package',
            'cpn_ref_price',
            'cpn_period',
            'cpn_detect_apartment',
            'status as cpn_status',
            'cpn_control_level',
            'equip_number',
            'yield_is_core_important as cpn_is_core_important',
            'safe_color',
            'result_grc as result',
        ];

        if ($cpn_type == 1)
        {
            $where .= " is_repeat <> 1 ";
            $where .= " and result_grc <> 1 ";

            $select1 = [
                'pro_review_opinion_add.cpn_category_code',
                'pro_review_opinion_add.cpn_manufacturer_replace as cpn_manufacturer',
                'pro_review_opinion_add.cpn_specification_model_replace as cpn_specification_model',
                'pro_review_opinion_add.cpn_name',
                'pro_review_opinion_add.cpn_quality',
                'pro_review_opinion_add.cpn_package',
                'pro_review_opinion_add.cpn_ref_price',
                'pro_review_opinion_add.cpn_period',
                'pro_review_opinion_add.cpn_detect_apartment',
                'pro_review_opinion_add.cpn_status',
                'cpn_import.cpn_control_level as cpn_control_level',
                'cpn_import.yield_is_core_important as cpn_is_core_important',
                'cpn_import.yield_safe_color as safe_color',
                'cpn_import.equip_number',
                'pro_review_opinion_import.pro_conclusion as result',
            ];

            $addQuery = ProReviewOpinionAdd::leftJoin('pro_review_opinion_import', 'pro_review_opinion_add.hash_code', '=', 'pro_review_opinion_import.hash_code')
                ->leftJoin('cpn_import', 'pro_review_opinion_import.cpn_id', '=', 'cpn_import.id')
                ->where(['pro_review_opinion_add.main_task_id' => $main_task_id, 'pro_review_opinion_add.type' => 1])
                ->select($select1);

            $cpnList = CpnDomestic::whereRaw($where)->select($select2)->union($addQuery)->get()->toArray();

            $cpnListGroup = CpnDomestic::whereRaw($where)->select($select2)->groupBy('cpn_manufacturer', 'cpn_specification_model')->union($addQuery)->count();

            $where2      = $where . 'and cpn_domestic.yield_is_core_important = 1';
            $cpnListCore = CpnDomestic::whereRaw($where2)->select($select2)->union($addQuery)->get()->toArray();

            $where3           = $where . 'and cpn_domestic.yield_is_core_important = 1';
            $cpnListGroupCore = CpnDomestic::whereRaw($where3)->select($select2)->groupBy('cpn_manufacturer', 'cpn_specification_model')->union($addQuery)->count();

            foreach ($cpnList as $key => $value)
            {
                if ($value['result'] == '国产替代纳入国产清单')
                {
                    $cpnList[$key]['result'] = 4;
                }
            }
        } elseif ($cpn_type == 2)
        {
            array_push($select2, 'yield_necessity');
            array_push($select2, 'dependence');

            $where   .= " is_repeat <> 1 ";
            $where   .= " and result_grc <> 1 ";
            $where   .= " and result_unite in (2,3,4) ";
            $cpnList = CpnImport::whereRaw($where)->select($select2)->get()->toArray();

            $cpnListGroup = CpnImport::whereRaw($where)->groupBy('cpn_specification_model')->select($select2)->count();

            $where2      = $where . 'and cpn_import.yield_is_core_important = 1';
            $cpnListCore = CpnImport::whereRaw($where2)->select($select2)->get()->toArray();

            $where3           = $where . 'and cpn_import.yield_is_core_important = 1';
            $cpnListGroupCore = CpnImport::whereRaw($where3)->select($select2)->groupBy('cpn_specification_model')->count();
        }

        return [
            'data'       => $cpnList,
            'count'      => $cpnListGroup,
            'core_data'  => $cpnListCore,
            'core_count' => $cpnListGroupCore,
        ];
    }

    /**
     * 统计数据
     * @param $cpnList
     * @return array
     */
    public function count3($cpnList, $type, $main_task_id)
    {
        $cpnData     = $cpnList['data'];
        $count       = $cpnList['count'];
        $cpnDataCore = $cpnList['core_data'];
        $countCore   = $cpnList['core_count'];

        $excel     = new Excel();
        $weightMap = $excel->getListWeightMap($main_task_id);

        if ($type == 1)
        {
            $arr3 = [
                ['level' => 'A', 'all' => 0, 'core' => 0],
                ['level' => 'B', 'all' => 0, 'core' => 0],
                ['level' => 'C', 'all' => 0, 'core' => 0],
                ['level' => 'D', 'all' => 0, 'core' => 0],
                ['level' => 'E', 'all' => 0, 'core' => 0],
            ];

            foreach ($weightMap as $mapKey => $mapValue)
            {
                foreach ($cpnData as $cpnKey => $cpnValue)
                {
                    $cpnData[$cpnKey]['num'] = $mapValue['weight'] * $cpnValue['equip_number'];

                    if ($cpnValue['cpn_control_level'] == 'A')
                    {
                        $arr3[0]['all']++;
                        if ($cpnValue['cpn_is_core_important'] == 1)
                        {
                            $arr3[0]['core']++;
                        }
                    } elseif ($cpnValue['cpn_control_level'] == 'B')
                    {
                        $arr3[1]['all']++;
                        if ($cpnValue['cpn_is_core_important'] == 1)
                        {
                            $arr3[1]['core']++;
                        }
                    } elseif ($cpnValue['cpn_control_level'] == 'C')
                    {
                        $arr3[2]['all']++;
                        if ($cpnValue['cpn_is_core_important'] == 1)
                        {
                            $arr3[2]['core']++;
                        }
                    } elseif ($cpnValue['cpn_control_level'] == 'D')
                    {
                        $arr3[3]['all']++;
                        if ($cpnValue['cpn_is_core_important'] == 1)
                        {
                            $arr3[3]['core']++;
                        }
                    } elseif ($cpnValue['cpn_control_level'] == 'E')
                    {
                        $arr3[4]['all']++;
                        if ($cpnValue['cpn_is_core_important'] == 1)
                        {
                            $arr3[4]['core']++;
                        }
                    }
                }

                foreach ($cpnDataCore as $cpnKeyCore => $cpnValueCore)
                {
                    $cpnDataCore[$cpnKeyCore]['num'] = $mapValue['weight'] * $cpnValueCore['equip_number'];
                }
            }


            $arr = [
                'list'  => [
                    'count' => count($cpnData),
                    'type'  => $count,
                    'num'   => array_sum(array_column($cpnData, 'num')),
                ],
                'core'  => [
                    'count' => count($cpnDataCore),
                    'type'  => $countCore,
                    'num'   => array_sum(array_column($cpnDataCore, 'num')),
                ],
                'level' => $arr3,
            ];

            return $arr;
        } elseif ($type == 2)
        {
            $arr3 = [
                ['color' => '红色', 'all' => 0, 'core' => 0],
                ['color' => '紫色', 'all' => 0, 'core' => 0],
                ['color' => '橙色', 'all' => 0, 'core' => 0],
                ['color' => '黄色', 'all' => 0, 'core' => 0],
                ['color' => '绿色', 'all' => 0, 'core' => 0],
            ];

            $arr4 = [
                ['level' => 'A', 'all' => 0, 'core' => 0],
                ['level' => 'B', 'all' => 0, 'core' => 0],
                ['level' => 'C', 'all' => 0, 'core' => 0],
            ];

            foreach ($weightMap as $mapKey => $mapValue)
            {
                foreach ($cpnData as $cpnKey => $cpnValue)
                {
                    $cpnData[$cpnKey]['num'] = $mapValue['weight'] * $cpnValue['equip_number'];

                    if ($cpnValue['safe_color'] == '红色')
                    {
                        $arr3[0]['all']++;
                        if ($cpnValue['cpn_is_core_important'] == 1)
                        {
                            $arr3[0]['core']++;
                        }
                    } elseif ($cpnValue['safe_color'] == '紫色')
                    {
                        $arr3[1]['all']++;
                        if ($cpnValue['cpn_is_core_important'] == 1)
                        {
                            $arr3[1]['core']++;
                        }
                    } elseif ($cpnValue['safe_color'] == '橙色')
                    {
                        $arr3[2]['all']++;
                        if ($cpnValue['cpn_is_core_important'] == 1)
                        {
                            $arr3[2]['core']++;
                        }
                    } elseif ($cpnValue['safe_color'] == '黄色')
                    {
                        $arr3[3]['all']++;
                        if ($cpnValue['cpn_is_core_important'] == 1)
                        {
                            $arr3[3]['core']++;
                        }
                    } elseif ($cpnValue['safe_color'] == '绿色')
                    {
                        $arr3[4]['all']++;
                        if ($cpnValue['cpn_is_core_important'] == 1)
                        {
                            $arr3[4]['core']++;
                        }
                    }

                    if ($cpnValue['dependence'] == '一级')
                    {
                        $arr4[0]['all']++;
                        if ($cpnValue['cpn_is_core_important'] == 1)
                        {
                            $arr4[0]['core']++;
                        }
                    } elseif ($cpnValue['dependence'] == '二级')
                    {
                        $arr4[1]['all']++;
                        if ($cpnValue['cpn_is_core_important'] == 1)
                        {
                            $arr4[1]['core']++;
                        }
                    } elseif ($cpnValue['dependence'] == '三级')
                    {
                        $arr4[2]['all']++;
                        if ($cpnValue['cpn_is_core_important'] == 1)
                        {
                            $arr4[2]['core']++;
                        }
                    }
                }

                foreach ($cpnDataCore as $cpnKeyCore => $cpnValueCore)
                {
                    $cpnDataCore[$cpnKeyCore]['num'] = $mapValue['weight'] * $cpnValueCore['equip_number'];
                }
            }


            $arr = [
                'list'  => [
                    'count' => count($cpnData),
                    'type'  => $count,
                    'num'   => array_sum(array_column($cpnData, 'num')),
                ],
                'core'  => [
                    'count' => count($cpnDataCore),
                    'type'  => $countCore,
                    'num'   => array_sum(array_column($cpnDataCore, 'num')),
                ],
                'color' => $arr3,
                'level' => $arr4,
            ];

            return $arr;
        }

    }

    /**
     * 获取清单数据
     * @param $main_task_id
     * @param $cpn_type
     * @param $list_id
     * @return array
     */
    public function getDoubleList($main_task_id, $cpn_type, $list_id)
    {
        $cpn_type = $cpn_type == '国产' ? 1 : 2;

        $where = '';

        if ($cpn_type == 1)
        {
            $where .= " is_repeat <> 1 ";
            $where .= " and result_grc <> 1 ";

            if (!empty($list_id))
                $where .= " and list_id = {$list_id} ";

            $addData = ProReviewOpinionAdd::with(['ProReviewOpinionImport'=>function($query){
                $query->with('cpn_import');
            }])
                ->where(['pro_review_opinion_add.main_task_id' => $main_task_id, 'pro_review_opinion_add.type' => 1])
                ->get()->toArray();

            $cpnList = CpnDomestic::whereRaw($where)->get()->toArray();

            $addArr = [];

            foreach ($addData as $key => $value)
            {
                $arr['id']                       = '';
                $arr['list_id']                  = $value['pro_review_opinion_import']['cpn_import']['list_id'];
                $arr['cpn_category_code']        = $value['cpn_category_code'];
                $arr['amount']                   = $value['pro_review_opinion_import']['cpn_import']['amount'];
                $arr['unique_code']              = '';
                $arr['cpn_ref_price']            = $value['cpn_ref_price'];
                $arr['cpn_name']                 = $value['cpn_name'];
                $arr['cpn_specification_model']  = $value['cpn_specification_model_replace'];
                $arr['cpn_quality']              = $value['cpn_quality'];
                $arr['cpn_package']              = $value['cpn_package'];
                $arr['cpn_country']              = $value['cpn_country'];
                $arr['cpn_manufacturer']         = $value['cpn_manufacturer_replace'];
                $arr['cpn_period']               = $value['cpn_period'];
                $arr['cpn_detect_apartment']     = $value['cpn_detect_apartment'];
                $arr['history']                  = '';
                $arr['dependence']               = '';
                $arr['safe_color']               = $value['safe_color'];
                $arr['proposed_safe_color']      = $value['proposed_safe_color'];
                $arr['necessity']                = $value['necessity'];
                $arr['access_channel']           = $value['access_channel'];
                $arr['equip_name']               = $value['pro_review_opinion_import']['cpn_import']['equip_name'];
                $arr['equip_research_apartment'] = $value['pro_review_opinion_import']['cpn_import']['equip_research_apartment'];
                $arr['equip_use_number']         = $value['pro_review_opinion_import']['cpn_import']['equip_use_number'];
                $arr['equip_number']             = $value['pro_review_opinion_import']['cpn_import']['equip_number'];
                $arr['remark']                   = '';
                $arr['status']                   = $value['cpn_status'];
                $arr['is_regular']               = 1;
                $arr['result_grc']               = 4;
                $arr['cpn_control_level']        = $value['pro_review_opinion_import']['cpn_import']['cpn_control_level'];
                $arr['yield_is_core_important']  = $value['pro_review_opinion_import']['cpn_import']['yield_is_core_important'];
                $arr['yield_control_level']      = $value['pro_review_opinion_import']['cpn_import']['cpn_control_level'];
                $arr['is_repeat']                = 0;
                array_push($addArr, $arr);
            }

            $cpnList = array_merge($addArr,$cpnList);

            foreach ($cpnList as $key => $value)
            {
                if ($value['result_grc'] == '国产替代纳入国产清单')
                {
                    $cpnList[$key]['result_grc'] = 4;
                }
            }


        } elseif ($cpn_type == 2)
        {
            $where .= " is_repeat <> 1 ";
            $where .= " and result_grc <> 1 ";
            $where .= " and result_unite in (2,3,4) ";

            if (!empty($list_id))
                $where .= " and list_id = {$list_id} ";

            $cpnList = CpnImport::whereRaw($where)->get()->toArray();
        }

        return [
            'data' => $cpnList,
        ];
    }

    public function getReportDomList($main_task_id)
    {

        $reArr = [
            'list'=>[
                'all'=>0,
                'pass'=>0,
                'noPass'=>0,
                'noPass01'=>0,
                'noPass02'=>0,
            ],
            'aux'=>[
                'coreType'=>0,
                'coreTypeP'=>'10%',
                'cCoreType'=>0,
                'cCoreTypeP'=>'10%',
                'coreList'=>[]
            ],
            'empty'=>[
                'type'=>0,
            ]
        ];


    }
}