<?php

namespace App\Http\Controllers;

use App\Model\AuxiliaryConfig;
use App\Model\CpnDomestic;
use App\Model\CpnFiles;
use App\Model\CpnImport;
use App\Model\Equipment;
use App\Model\Excel;
use App\Model\MainTask;
use App\Model\MainTaskStage;
use App\Model\ModelStructure;
use App\Model\ProNew;
use App\Model\ProReviewError;
use App\Model\ProReviewOpinionAdd;
use App\Model\ProReviewOpinionImport;
use App\Model\Report;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpParser\Node\Scalar\MagicConst\Dir;

class ProNewController extends Controller
{
    /**
     * 专家主页展示
     * @param Request $request
     * @return \Illuminate\View\View
     */
    public function index(Request $request)
    {
        $page_size = $request->get('pages', 10);
        $typePage  = $request->get('typePage', 0);

        $data = MainTask::with(['main_stage', 'equipment_model'])
            ->whereIn('status', [1, 2])
            ->where(['professor_status' => $typePage])
            ->orderBy('id', 'desc')
            ->orderBy('professor_status_ext', 'desc')
            ->paginate($page_size);

        foreach ($data as $key => $value)
        {
            $modelData = Equipment::where('id', $value['model_id'])->first();
            // 总师单位
            $value->company = @$modelData->company;
            $value->name    = @$modelData->name;
        }

        $res = [
            // 清单型号任务信息
            'data' => $data->toArray(),
        ];

        return $this->success($res);
    }

    /**
     * 获取组织审查需要的数据
     * @param Request $request
     * @return \Illuminate\Http\JsonResponse
     */
    public function distribution(Request $request)
    {
        $main_task_id = $request->get('main_task_id', null);
        if (empty($main_task_id))
        {
            return $this->error('参数不能为空');
        }


        $mainTaskData = MainTask::find($main_task_id);//主任务信息

        if (empty($mainTaskData))
        {
            return $this->error('信息不存在');
        }

        $data = [
            'id'                       => $main_task_id,
            'equipment_name'           => @$mainTaskData->equipment_model->name,
            'stage_name'               => @$mainTaskData->main_stage->name,
            'equipment_competent_unit' => @$mainTaskData->equipment_competent_unit,
            'company'                  => @$mainTaskData->equipment_model->company,
            'edition_number'           => @$mainTaskData->edition_number,
            'time'                     => @$mainTaskData->start_time,
        ];

        return $this->success($data);
    }

    /**
     * 导出专家审查数据包
     * @param Request $request
     * @return \Illuminate\Http\JsonResponse
     */
    public function getDataPackage(Request $request)
    {
        $main_task_id = $request->get('main_task_id', null);
        $unit_one     = $request->get('unit_one', null);
        $unit_two     = $request->get('unit_two', null);
        $time         = $request->get('time', null);

        if (empty($main_task_id))
            return $this->error('参数不能为空');
        if (empty($unit_one))
            return $this->error('参数不能为空');
        if (empty($unit_two))
            return $this->error('参数不能为空');
        if (empty($time))
            return $this->error('参数不能为空');

        $mainTaskData = MainTask::with('equipment_model')->with('main_stage')->find($main_task_id);

        if (empty($mainTaskData))
            return $this->error('信息不存在');

        $equipment_name = $mainTaskData->equipment_model->name;
        $stage          = $mainTaskData->main_stage->name;

        // sqlite数据库兼容性
        $sqlType = env('DB_TYPE', 'mysql');

        if ($sqlType === 'sqlite')
            $insertSize = 30;
        else
            $insertSize = 1000;

        DB::beginTransaction();
        try
        {
            // 获取并且处理要导出的数据
            $ProNew     = new ProNew();
            $data       = $ProNew->getExportData($main_task_id);
            $importData = $ProNew->getInstallData($main_task_id);

            // 定义文件名称和文件夹名称
            $excelDir = 'attachment/package/' . date('Y-m-d') . '/' . sha1(time()) . '/';
            $fileName = $equipment_name . '型号' . $stage . '阶段用研结合';

            if (!file_exists($excelDir))
            {
                $makeResult = mkdir($excelDir, 0777, true);

                if (!$makeResult)
                    return $this->error('err 001');
            }

            // 当元器件list为不为空的时候
            if (!empty($importData))
            {

                $chunks = collect($importData)->chunk($insertSize);
                foreach ($chunks as $block)
                {
                    ProReviewOpinionImport::insert($block->toArray());
                }

                $result = $ProNew->export($data, $excelDir, $fileName);

                if ($result['status'] === true)
                {
                    $update = [
                        'requester_unit'            => $unit_one,
                        'review_unit'               => $unit_two,
                        'professor_expect_end_time' => $time,
                        'professor_status'          => 1,
                        'professor_status_ext'      => 1
                    ];

                    $res = MainTask::where('id', $main_task_id)->update($update);

                    if (1)
                    {
                        DB::commit();
                        return $this->success($result, '导出成功');
                    }

                    DB::rollBack();
                    return $this->error('err 003');

                } else
                {
                    DB::rollBack();
                }
                return $this->error('err 002');
            } else
            {
                DB::rollBack();
                return $this->error('无可用元器件');
            }

        } catch (\Exception $exception)
        {
            DB::rollBack();
            return $this->error($exception, 'err');
        }
    }

    /**
     * 导入专家审查数据包
     * @param Request $request
     * @return \Illuminate\Http\JsonResponse
     * @throws \Illuminate\Contracts\Container\BindingResolutionException
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     */
    public function exportPackage(Request $request)
    {
        $rootPath     = public_path('attachment/unPackage');
        $main_task_id = $request->input('main_task_id', null);

        if (empty($main_task_id))
            return $this->error('任务ID不能为空');

        $mainTaskData = MainTask::with('equipment_model')->with('main_stage')->find($main_task_id);

        if (empty($mainTaskData))
            return $this->error('信息不存在');

        $equipment_name = $mainTaskData->equipment_model->name;
        $stage          = $mainTaskData->main_stage->name;

        $fileName = $equipment_name . '型号' . $stage . '阶段用研结合';

        if ($request->hasFile('package'))
        {
            $file      = $request->file('package');
            $extension = $file->getClientOriginalExtension();
            $file_name = $file->getClientOriginalName();

            $file_first = current(explode('.', $file_name));

            if ($extension != 'zip')
                return $this->error('上传文件格式不正确');

            set_time_limit(0);
            ini_set('memory_limit', '-1');

            $proNew = new ProNew();
            $excel  = new Excel();

            // 解压文件
            $unzipResult = $proNew->unzip($rootPath, $file);
            if ($unzipResult['status'] !== true)
                return $this->error($unzipResult['data']);

            // 判断文件是否是导出的文件
            $verifyResult = $proNew->verifyFile($unzipResult['data'], $fileName);
            if ($verifyResult['status'] !== true)
                return $this->error($verifyResult['msg']);

            // 解压后文件存放目录
            $fileDir = $rootPath . DIRECTORY_SEPARATOR . $file_first . DIRECTORY_SEPARATOR . 'package' . DIRECTORY_SEPARATOR;

            // 获取要读取的文件名称
            $excelName_01 = $proNew->getFileName($unzipResult['data'], $fileName, '审查表');
            $excelName_02 = $proNew->getFileName($unzipResult['data'], $fileName, '审查报告');

            // 读取文件
            $readExcelResult01 = $excel->readExcel($fileDir . $excelName_01['msg']);
            $readExcelResult02 = $excel->readExcel($fileDir . $excelName_02['msg'], 20, 2, [0, 1]);

            // 验证文件是否可读
            $verifyRes01 = $proNew->verifyExcelData01($readExcelResult01);
            if ($verifyRes01['status'] !== true)
                return $this->error($verifyRes01['msg']);

            $verifyRes02 = $proNew->verifyExcelData02($readExcelResult02);
            if ($verifyRes02['status'] !== true)
                return $this->error($verifyRes02['msg']);

            // 导入数据
            $exportResult = $proNew->exportExcelData($main_task_id, $readExcelResult01, $readExcelResult02);
            if ($exportResult['status'] === true)
                return $this->success([], $exportResult['msg']);
            else
                return $this->error($exportResult['msg'], []);
        }
        return $this->error([], '无文件上传');
    }

    /**
     * 删除已经导入的数据包
     * @param Request $request
     * @return \Illuminate\Http\JsonResponse
     */
    public function deleteExportPackage(Request $request)
    {
        $main_task_id = $request->get('main_task_id', null);

        if (empty($main_task_id))
            return $this->error('参数不能为空');

        $mainTaskData = MainTask::find($main_task_id);

        if (empty($mainTaskData))
            return $this->error('信息不存在');

        $updateArr = [
            'kb_substitution_plan'    => '',
            'kb_substitution_model'   => '',
            'kb_substitution_mfr'     => '',
            'kb_substitution_whether' => '',
            'kb_substitution_status'  => '',
            'pro_massage'             => '',
            'pro_massage_status'      => '',
            'pro_way'                 => '',
            'pro_conclusion'          => '',
        ];

        DB::beginTransaction();
        try
        {
            ProReviewOpinionImport::where('main_task_id', $main_task_id)->update($updateArr);
            ProReviewOpinionAdd::where('main_task_id', $main_task_id)->delete();
            MainTask::where('main_task_id', $main_task_id)->update(['professor_status_ext' => 0]);

            DB::commit();
            return $this->success([], '删除成功');
        } catch (\Exception $exception)
        {
            DB::rollBack();
            return $this->error('删除失败');
        }
    }

    /**
     * 导出错误报告
     * @param Request $request
     * @return array
     */
    public function exportErrData(Request $request)
    {
        $main_task_id = $request->get('main_task_id', null);

        if (empty($main_task_id))
            return $this->error('参数不能为空');

        $mainTaskData = MainTask::with('equipment_model')->with('main_stage')->find($main_task_id);

        if (empty($mainTaskData))
            return $this->error('信息不存在');

        $equipment_name = $mainTaskData->equipment_model->name;
        $stage          = $mainTaskData->main_stage->name;

        // 获取错误数据
        $errData = ProReviewError::where('main_task_id', $main_task_id)->get()->toArray();

        if (empty($errData))
            return $this->success([], '无错误信息');

        $proNew = new ProNew();

        $fileName = $equipment_name . '型号' . $stage . '阶段用研结合';
        $excelDir = 'attachment/error/';

        // 生成错误的excel
        $exportRes = $proNew->exportError($errData, $excelDir, $fileName);

        if ($exportRes['status'] === true)
            return $this->success($exportRes, '按照文件路径下载即可');
        else
            return $this->error($exportRes['msg'], $exportRes);
    }

    /**
     * 清单结论 统计
     * @param Request $request
     * @return \Illuminate\Http\JsonResponse
     */
    public function look1(Request $request)
    {
        $main_task_id = $request->get('main_task_id', null);

        if (empty($main_task_id))
            return $this->error('参数不能为空');

        $mainTaskData = MainTask::find($main_task_id);

        if (empty($mainTaskData))
            return $this->error('信息不存在');

        $arr = ProReviewOpinionImport::where('main_task_id', $main_task_id)
            ->whereIn('pro_conclusion', ['国产替代纳入国产清单', '研制攻关暂纳入进口清单', '纳入进口清单', '替换其他规格型号纳入进口清单', '不选用'])
            ->get()
            ->groupBy('pro_conclusion')
            ->toArray();

        $all = 0;

        foreach ($arr as $key => $value)
        {
            $all += count($value);
        }

        $returnArr = [
            'data' => [
                '国产替代纳入国产清单'     => @count(@$arr['国产替代纳入国产清单']),
                '研制攻关暂纳入进口清单'    => @count(@$arr['研制攻关暂纳入进口清单']),
                '纳入进口清单'         => @count(@$arr['纳入进口清单']),
                '替换其他规格型号纳入进口清单' => @count(@$arr['替换其他规格型号纳入进口清单']),
                '不选用'            => @count(@$arr['不选用'])
            ],

            'all' => $all
        ];

        return $this->success($returnArr, 'ok');
    }

    /**
     * 查看清单详情
     * @param Request $request
     * @return \Illuminate\Http\JsonResponse
     */
    public function look2(Request $request)
    {
        $main_task_id = $request->get('main_task_id', null);
        $page_size    = $request->get('pages', null);
        $type         = $request->get('type', null);

        $arr = [
            1 => '国产替代纳入国产清单',
            2 => '研制攻关暂纳入进口清单',
            3 => '纳入进口清单',
            4 => '替换其他规格型号纳入进口清单',
            5 => '不选用',
        ];

        if (empty($main_task_id))
            return $this->error('参数不能为空');

        if (!in_array($type, array_keys($arr)))
            return $this->error('参数错误');

        $mainTaskData = MainTask::find($main_task_id);

        if (empty($mainTaskData))
            return $this->error('信息不存在');

        $arr = ProReviewOpinionImport::where('main_task_id', $main_task_id)
            ->with(['cpn_import' => function ($query) {
                $query->select(
                    'cpn_import.id',
                    'cpn_import.cpn_specification_model',
                    'cpn_import.cpn_manufacturer',
                    'cpn_import.cpn_quality',
                    'cpn_import.cpn_package',
                    'cpn_import.cpn_control_level',
                    'cpn_import.yield_is_core_important',
                    'cpn_import.yield_safe_color',
                    'cpn_import.result_pc'
                );
            }])
            ->select(
                'pro_review_opinion_import.main_task_id',
                'pro_review_opinion_import.cpn_id',
                'pro_review_opinion_import.kb_substitution_plan',
                'pro_review_opinion_import.kb_substitution_model',
                'pro_review_opinion_import.kb_substitution_mfr',
                'pro_review_opinion_import.kb_substitution_whether',
                'pro_review_opinion_import.kb_substitution_status',
                'pro_review_opinion_import.pro_massage',
                'pro_review_opinion_import.pro_massage_status',
                'pro_review_opinion_import.pro_way',
                'pro_review_opinion_import.pro_conclusion'
            )
            ->with('AuxiliaryResultReplace')
            ->where('pro_conclusion', $arr[$type])
            ->paginate($page_size)
            ->toArray();

        foreach ($arr['data'] as $key => $value)
        {
            $kb_aux_substitution_model  = '';
            $kb_aux_substitution_status = '';
            if (!empty($value['auxiliary_result_replace']))
            {
                foreach ($value['auxiliary_result_replace'] as $reKey => $reValue)
                {
                    $kb_aux_substitution_model           .= $reValue['replace_cpn_specification_model'] . '[' . $reValue['replace_cpn_manufacturer'] . '|' . $reValue['replace_cpn_quality'] . '];';
                    $kb_aux_substitution_status          .= $reValue['replace_product_state'];
                    $arr['data'][$key]['replace_str']    = $kb_aux_substitution_model;
                    $arr['data'][$key]['replace_status'] = $kb_aux_substitution_status;
                }
            } else
            {
                $arr['data'][$key]['replace_str']    = '';
                $arr['data'][$key]['replace_status'] = '';
            }
        }

        return $this->success($arr, 'ok');
    }

    /**
     * 专家导出报告
     * @param Request $request
     * @return \Illuminate\Http\JsonResponse
     */
    public function makeDoc(Request $request)
    {
        $main_task_id = $request->get('main_task_id', null);

        if (empty($main_task_id))
            return $this->error('参数不能为空');

        $mainTask  = MainTask::find($main_task_id);
        $modelName = Equipment::find($mainTask['model_id']);
        $stageName = MainTaskStage::find($mainTask['stage_id']);

        // 国产数据
        $request['main_task_id'] = $main_task_id;
        $request['type']         = 1;
        $cpn_data_dom            = $this->look3($request);
        $cpn_data_dom            = $cpn_data_dom->original;
        $cpn_data_dom            = $cpn_data_dom['data'];

        // 进口数据
        $request['main_task_id'] = $main_task_id;
        $request['type']         = 2;
        $cpn_data_imp            = $this->look3($request);
        $cpn_data_imp            = $cpn_data_imp->original;
        $cpn_data_imp            = $cpn_data_imp['data'];

        $proNew          = new ProNew();
        $cpn_data_dom_02 = $proNew->getReportDomList($main_task_id);
        dd($cpn_data_dom_02);

        $data = json_decode(file_get_contents(storage_path('wei.json')), true);

        $report  = new Report();
        $phpWord = new PhpWord();

        // 页边距
        $sectionStyle = [
            'orientation'  => null,
            'marginLeft'   => 1701,
            'marginRight'  => 1418,
            'marginTop'    => 1418,
            'marginBottom' => 1418,
        ];

        $phpWord->setDefaultFontName('宋体');

        // 居中+行间距 前20磅 后20磅
        $center = ['align' => 'center', 'spaceAfter' => 20 * 20, 'spaceBefore' => 20 * 20];
        // 居中
        $alignCenter = ['align' => 'center'];
        // 行间距 前11磅 后11磅
        $spacing100 = ['spaceBefore' => 11 * 20, 'spaceAfter' => 11 * 20];
        // 行间距 前6磅 后6磅
        $spacing50 = ['spaceBefore' => 6 * 20, 'spaceAfter' => 6 * 20];
        // 标题样式
        $phpWord->addFontStyle('mainTitle', ['bold' => true, 'color' => 'black', 'size' => 16, 'name' => '宋体']);
        // 正文样式
        $phpWord->addFontStyle('desc', ['bold' => false, 'color' => 'black', 'size' => 13, 'name' => '宋体']);
        // 小标题样式
        $phpWord->addFontStyle('littleTitle', ['bold' => true, 'color' => 'black', 'size' => 14, 'name' => '宋体']);
        // 表格表头样式
        $phpWord->addFontStyle('tableTitle', ['bold' => true, 'color' => 'black', 'size' => 10.5, 'name' => '宋体']);
        // 表格正文样式/注释样式
        $phpWord->addFontStyle('tableDesc', ['bold' => false, 'color' => 'black', 'size' => 10.5, 'name' => '宋体']);
        // 表格样式
        $phpWord->addTableStyle('table', array('borderSize' => 6, 'borderColor' => 'black', 'cellMargin' => 80, 'align' => 'center'));
        $phpWord->addTableStyle('tableRight', array('borderSize' => 6, 'borderColor' => 'black', 'cellMargin' => 80, 'align' => 'right'));

        // 四号
        $phpWord->setDefaultFontSize(14);

        // 封面表格样式
        $phpWord->addTableStyle('feng_table', array('bold' => true, 'borderSize' => 0, 'borderColor' => 'white', 'cellMargin' => 80, 'align' => 'center'));

        $section = $phpWord->createSection($sectionStyle);

        // 页脚
        $footer = $section->createFooter();
        $footer->addPreserveText('第 {PAGE} / {NUMPAGES} 页', 'desc', ['align' => 'center']);

        // 四号
        $phpWord->setDefaultFontSize(14);

        //封面
        $table = $section->addTable('tableRight');
        $table->addRow(25);
        $table->addCell(1300)->addText('密 级：', 'mainTitle');
        $table->addCell(2300)->addText('', 'mainTitle');
        $table->addRow(25);
        $table->addCell(1300)->addText('编 号：', 'mainTitle');
        $table->addCell(2300)->addText($data['number'], 'mainTitle');
        $table->addRow(25);
        $table->addCell(1300)->addText('页 数：', 'mainTitle');
        $table->addCell(2300)->addText('', 'mainTitle');

        $section->addTextBreak(4);

        $section->addText('《关于加强武器装备使用国产电子元器件管理的意见》双清单审查报告', ['bold' => true, 'size' => 12], ['align' => 'center']);
        $section->addText($modelName->model_name . '装备型号' . $stageName->stage . '阶段<w:br/>国产和进口电子元器件清单<w:br/>审查报告', ['bold' => true, 'size' => 28], ['align' => 'center']);

        $section->addTextBreak(7);

        $table = $section->addTable('feng_table');
        $table->addRow(25);
        $table->addCell(2500)->addText('主建部门：', ['bold' => true, 'size' => 16]);
        $table->addCell(3000, ['borderBottomColor' => 'black', 'borderBottomSize' => 6])->addText('');
        $table->addRow(25);
        $table->addCell(2500)->addText('审查机构：', ['bold' => true, 'size' => 16]);
        $table->addCell(3000, ['borderBottomColor' => 'black', 'borderBottomSize' => 6])->addText('');
        $table->addRow(25);
        $table->addCell(2500)->addText('主    审：', ['bold' => true, 'size' => 16]);
        $table->addCell(3000, ['borderBottomColor' => 'black', 'borderBottomSize' => 6])->addText('');
        $table->addRow(25);
        $table->addCell(2500)->addText('审    核：', ['bold' => true, 'size' => 16]);
        $table->addCell(3000, ['borderBottomColor' => 'black', 'borderBottomSize' => 6])->addText('');
        $table->addRow(25);
        $table->addCell(2500)->addText('批    准：', ['bold' => true, 'size' => 16]);
        $table->addCell(3000, ['borderBottomColor' => 'black', 'borderBottomSize' => 6])->addText('');


        $section = $phpWord->createSection($sectionStyle);
        $section->addText('审查报告概述', 'mainTitle', $center);

        $section->addTitle('一、送审情况', "littleTitle", $spacing100);

        $section->addText("编制单位：    xx所", 'desc', $spacing50);
        $section->addText("研制阶段：    " . $stageName->stage . "阶段", 'desc', $spacing50);
        $section->addText("双清单版本号：军兵种编号+型号名称首字母+阶段编号+审查序号", 'desc', $spacing50);
        $section->addText("委托审查单位：{$mainTask->review_unit} 单位", 'desc', $spacing50);
        $section->addText("送审时间：    " . $mainTask->start_time, 'desc', $spacing50);
        $section->addText("软件版本号：  武器装备电子元器件管理信息系统V1.0");
        $section->addText("审查支撑数据：■国产军用电子元器件手册-2019版（简称《国产手册》）", 'desc', $spacing50);
        $section->addText("             ■进口电子元器件控制等级信息清单-2019版", 'desc', $spacing50);
        $section->addText("             ■国产军用关键软硬件推荐产品名录");
        $section->addText("             ☑第三方数据：赛思库®数据-2020版（简称《赛思库》）", 'desc', $spacing50);
        $section->addText("             □自主可控基础产品目录", 'desc', $spacing50);
        $section->addText("             □对外依存度数据", 'desc', $spacing50);
        $section->addText("             □其他受控数据源：XXXXXXX", 'desc', $spacing50);

        // 二、送审清单内容摘要
        $section = $phpWord->createSection($sectionStyle);

        $section->addText('二、送审清单内容摘要', 'littleTitle', $spacing100);
        $section->addText('1、国产电子元器件清单', 'littleTitle', $spacing100);
        $section->addText('（1）清单共有' . $cpn_data_dom['list']['count'] . '条数据，规格' . $cpn_data_dom['list']['type'] . '种，数量' . $cpn_data_dom['list']['num'] . '只', 'desc', $spacing50);
        $section->addText('（2）核心关键共有' . $cpn_data_dom['core']['count'] . '条数据，规格' . $cpn_data_dom['core']['type'] . '种，数量' . $cpn_data_dom['core']['num'] . '只', 'desc', $spacing50);
        $section->addText('（3）自主可控等级：', 'desc', $spacing50);
        $table01 = $section->addTable('table');
        $table01->addRow(20);
        $table01->addCell(3000)->addText('自主可控等级', 'tableTitle', $alignCenter);
        $table01->addCell(3000)->addText('全部', 'tableTitle', $alignCenter);
        $table01->addCell(3000)->addText('核心关键', 'tableTitle', $alignCenter);
        foreach ($cpn_data_dom['level'] as $key => $value)
        {
            $table01->addRow(20);
            $table01->addCell(3000)->addText($value['level'], 'tableDesc', $alignCenter);
            $table01->addCell(3000)->addText($value['all'], 'tableDesc', $alignCenter);
            $table01->addCell(3000)->addText($value['core'], 'tableDesc', $alignCenter);
        }

        $section->addText('2、进口电子元器件清单', 'littleTitle', $spacing100);
        $section->addText('（1）清单共有' . $cpn_data_imp['list']['count'] . '条数据，规格' . $cpn_data_imp['list']['type'] . '种，数量' . $cpn_data_imp['list']['num'] . '只', 'desc', $spacing50);
        $section->addText('（2）核心关键共有' . $cpn_data_imp['core']['count'] . '条数据，规格' . $cpn_data_imp['core']['type'] . '种，数量' . $cpn_data_imp['core']['num'] . '只', 'desc', $spacing50);
        $section->addText('（3）安全等级颜色：', 'desc', $spacing50);
        $table02 = $section->addTable('table');
        $table02->addRow(20);
        $table02->addCell(3000)->addText('安全等级颜色', 'tableTitle', $alignCenter);
        $table02->addCell(3000)->addText('全部', 'tableTitle', $alignCenter);
        $table02->addCell(3000)->addText('核心关键', 'tableTitle', $alignCenter);
        foreach ($cpn_data_imp['color'] as $key => $value)
        {
            $table02->addRow(20);
            $table02->addCell(3000)->addText($value['color'], 'tableDesc', $alignCenter);
            $table02->addCell(3000)->addText($value['all'], 'tableDesc', $alignCenter);
            $table02->addCell(3000)->addText($value['core'], 'tableDesc', $alignCenter);
        }

        $section = $phpWord->createSection($sectionStyle);
        $section->addText('审查结果-国产电子元器件清单', 'mainTitle', $center);

        $section->addTitle('一、合规性审查结论', "littleTitle", $spacing100);
        $section->addText("    国产电子元器件清单共计" . $data['regular_check_dom']['dom_total'] . "条，经审查有" . $data['regular_check_dom']['follow_regular'] . "条合规，" . $data['regular_check_dom']['no_follow_regular'] . "条不合规。", 'desc', $spacing50);
        $section->addText("    对不合规数据的处理：", 'desc', $spacing50);
        $section->addText("    （1）" . $data['regular_check_dom']['no_follow_regular_1'] . "条存在型号规格、生产厂商、分类代码缺失，或分类代码不属于GJB 8118的问题，判定审查不通过。", 'desc', $spacing50);
        $section->addText("    （2）" . $data['regular_check_dom']['no_follow_regular_2'] . "条按默认让步接收规则处理后进行审查。", 'desc', $spacing50);

        $section->addTitle('二、计算机辅助比对审查情况', "littleTitle", $spacing100);
        $section->addTitle('1、核心关键占比情况', "littleTitle", $spacing100);
        $section->addText("     （1）国产核心关键" . $data['regular_check_dom']['key_percent']['key_spec'] . "种，占总体" . $data['regular_check_dom']['key_percent']['total_percent'] . "。", 'desc', $spacing50);
        $section->addText("     （2）自主可控等级C级以上的国产核心关键" . $data['regular_check_dom']['key_percent']['control_key_spec'] . "种，占总体" . $data['regular_check_dom']['key_percent']['control_key_percent'], 'desc', $spacing50);
        $section->addText("     （3）国产核心关键种类分布：", 'desc', $spacing50);
        //生成饼状图
        $img = $report->generatePieGraph2($data['regular_check_dom']['key_spec_percent']);
        //图片路径插入到 word
        if ($img)
        {
            $section->addImage($img, ['width' => 450, 'height' => 250, 'ailgn' => 'center']);
        }

        $section->addTextBreak(1);

        $section->addTitle('2、自主可控等级分布情况', "littleTitle", $spacing100);
        $section->addText("     （1）A级：占比" . $data['regular_check_dom']['control_grade_a']['percent'] . "，核心关键占比" . $data['regular_check_dom']['control_grade_a']['key_percent']);
        $section->addText("     （2）B级：占比" . $data['regular_check_dom']['control_grade_b']['percent'] . "，核心关键占比" . $data['regular_check_dom']['control_grade_b']['key_percent']);
        $section->addText("     （1）C级：占比" . $data['regular_check_dom']['control_grade_c']['percent'] . "，核心关键占比" . $data['regular_check_dom']['control_grade_c']['key_percent']);
        $section->addText("     （1）D级：占比" . $data['regular_check_dom']['control_grade_d']['percent'] . "，核心关键占比" . $data['regular_check_dom']['control_grade_d']['key_percent']);
        $section->addText("     （1）E级：占比" . $data['regular_check_dom']['control_grade_e']['percent'] . "，核心关键占比" . $data['regular_check_dom']['control_grade_e']['key_percent']);

        $section->addTitle('3、伪空包抽查', "littleTitle", $spacing100);
        $section->addText("     待抽查，推荐抽查" . $data['regular_check_dom']['empty_spec'] . "种", 'desc', $spacing50);

        // 审查结果-进口电子元器件清单
        $section = $phpWord->createSection($sectionStyle);

        $section->addText('审查结果-进口电子元器件清单', 'mainTitle', $center);
        $section->addText('一、合规性审查结论', 'littleTitle', $spacing100);
        $section->addText('    进口电子元器件清单共计' . $data['regular_check_import']['dom_total'] . '条，经审查有' . $data['regular_check_import']['follow_regular'] . '条合规，' . $data['regular_check_import']['no_follow_regular'] . '条不合规，默认采取让步接收处理。（或要求研制单位修改核准后再上报。）', 'desc', $spacing50);
        $section->addText('    对不合规数据的处理：', 'desc', $spacing50);
        $section->addText('    （1）' . $data['regular_check_import']['no_follow_regular_1'] . '条存在型号规格、生成厂商、分类代码缺失，或分类代码不属于GJB 8118的问题，审查不通过。', 'desc', $spacing50);
        $section->addText('    （2）' . $data['regular_check_import']['no_follow_regular_2'] . '条按默认的异常处理规则进行数据处理后进行审查。', 'desc', $spacing50);

        $section->addText('二、计算机辅助比对审查结论', 'littleTitle', $spacing100);
        $section->addText('    进入计算机辅助比对审查的进口电器元器件数据共计' . $data['regular_check_import']['check_total'] . '条，规格' . $data['regular_check_import']['check_spec'] . '种。其中让步接受' . $data['regular_check_import']['check_recept'] . '条，规格' . $data['regular_check_import']['check_recept_spec'] . '种', 'desc', $spacing50);
        $section->addText('    YCD一级' . $data['regular_check_import']['ycd_one']['spec_total'] . '种，有' . $data['regular_check_import']['ycd_one']['yan_recept'] . '种让步接受的进入用研结合审查；  有安全风险' . $data['regular_check_import']['ycd_one']['manager_control'] . '种不允许选用；有' . $data['regular_check_import']['ycd_one']['safety_no_allow'] . '种纳入进口清单管控。', 'desc', $spacing50);
        $section->addText('    YCD二级' . $data['regular_check_import']['ycd_two']['spec_total'] . '种，其' . $data['regular_check_import']['ycd_two']['no_safety'] . '种无安全和可保障风险，且非核心关键器件的纳入进口清单管控；其余' . $data['regular_check_import']['ycd_two']['yong_yan'] . '种进入用研结合审查（含' . $data['regular_check_import']['ycd_two']['yan_recept'] . '种让步接收）', 'desc', $spacing50);
        $section->addText('    YCD三级' . $data['regular_check_import']['ycd_three']['spec_total'] . '种，有' . $data['regular_check_import']['ycd_three']['yan_recept'] . '种让步接收的进入用研结合审查；' . $data['regular_check_import']['ycd_three']['no_safety'] . '种无安全和可保障风险，且非核心关键器件的纳入进口清单管控；有' . $data['regular_check_import']['ycd_three']['dom_replace'] . '种国产替代产品价格过高、供货周期过长，纳入进口清单管控；其余' . $data['regular_check_import']['ycd_three']['other_dom_list'] . '种纳入国产清单。', 'desc', $spacing50);

        $section->addText('三、用研结合审查', 'littleTitle', $spacing100);
        $section->addText('    进入用研结合审查的进口电子元器件共计' . $data['regular_check_import']['yong_yan_check']['total'] . '种，经审查有' . $data['regular_check_import']['yong_yan_check']['pass_dom_replace'] . '种可通过系统级优化或研制攻关采用国产元器件替代，审查通过纳入国产清单；有' . $data['regular_check_import']['yong_yan_check']['no_pass_dom_replace'] . '种审查不通过纳入进口清单管控。', 'desc', $spacing50);

        $section->addText('四、结论', 'littleTitle', $spacing100);
        $section->addText('    ' . $data['regular_check_import']['bring_in_spec'] . '种纳入进口清单管控，' . $data['regular_check_import']['replace_spec_product'] . '种选用国产替代产品。' . $data['regular_check_import']['safety_no_allow'] . '种有安全风险不允许选用', 'desc', $spacing50);

        // 综合审查结论
        $section = $phpWord->createSection($sectionStyle);

        $section->addText('综合审查结论', 'mainTitle', $center);
        $section->addText('一、审查后国产电子元器件清单', 'littleTitle', $spacing100);
        $section->addText('（1）全部：规格' . $data['comprehensive_check_domestic']['total_spec'] . '种', 'desc', $spacing50);
        $section->addText('（2）核心关键：规格' . $data['comprehensive_check_domestic']['key_spec'] . '种', 'desc', $spacing50);
        $section->addText('（3）自主可控等级：', 'desc', $spacing50);
        $table01 = $section->addTable('table');
        $table01->addRow(20);
        $table01->addCell(3000)->addText('自主可控等级', 'tableTitle', $alignCenter);
        $table01->addCell(3000)->addText('全部', 'tableTitle', $alignCenter);
        $table01->addCell(3000)->addText('核心关键', 'tableTitle', $alignCenter);
        foreach ($data['comprehensive_check_domestic']['control_grade'] as $key => $value)
        {
            $table01->addRow(20);
            $table01->addCell(3000)->addText($value['name'], 'tableDesc', $alignCenter);
            $table01->addCell(3000)->addText($value['total'], 'tableDesc', $alignCenter);
            $table01->addCell(3000)->addText($value['key'], 'tableDesc', $alignCenter);
        }

        $section->addText('二、审查后进口电子元器件清单', 'littleTitle', $spacing100);
        $section->addText('（1）全部：规格' . $data['comprehensive_check_import']['total_spec'] . '种', 'desc', $spacing50);
        $section->addText('（2）核心关键：规格' . $data['comprehensive_check_import']['key_spec'] . '种', 'desc', $spacing50);
        $section->addText('（3）安全等级颜色：', 'desc', $spacing50);
        $table02 = $section->addTable('table');
        $table02->addRow(20);
        $table02->addCell(3000)->addText('安全等级颜色', 'tableTitle', $alignCenter);
        $table02->addCell(3000)->addText('全部', 'tableTitle', $alignCenter);
        $table02->addCell(3000)->addText('核心关键', 'tableTitle', $alignCenter);
        foreach ($data['comprehensive_check_import']['control_grade'] as $key => $value)
        {
            $table02->addRow(20);
            $table02->addCell(3000)->addText($value['name'], 'tableDesc', $alignCenter);
            $table02->addCell(3000)->addText($value['total'], 'tableDesc', $alignCenter);
            $table02->addCell(3000)->addText($value['key'], 'tableDesc', $alignCenter);
        }

        try
        {
            $fileName = '/attachment/report/' . ($data['model_name'] . '型号' . $data['stage'] . '阶段形式审查报告.docx');
            $writer   = IOFactory::createWriter($phpWord, 'Word2007');
            $writer->save('.' . $fileName);
            return $this->success($fileName);
        } catch (\Exception $e)
        {
            return $this->error($e->getMessage());
        }
    }

    /**
     * 清单结论 统计
     * @param Request $request
     * @return \Illuminate\Http\JsonResponse
     */
    public function look3(Request $request)
    {
        $main_task_id = $request->get('main_task_id', null);
        $type         = $request->get('type', null); // 1 国产 2 进口

        if (empty($main_task_id))
            return $this->error('参数不能为空');

        $mainTaskData = MainTask::find($main_task_id);

        if (empty($mainTaskData))
            return $this->error('信息不存在');

        $proNew = new ProNew();

        $cpnList = $proNew->getCpnData($main_task_id, $type);

        $countList = $proNew->count3($cpnList, $type, $main_task_id);

        return $this->success($countList, 'ok');
    }

    /**
     * 双清单 - 结构树
     * @param Request $request
     * @return mixed
     */
    public function getTree(Request $request)
    {
        $main_task_id = $request->get('main_task_id', null);
        if (empty($main_task_id))
            return $this->error('参数不能为空');

        if (!$mainTask = MainTask::find($main_task_id))
            $this->error('信息不存在');

        //元器件清单结构树数据
        $CpnFiles = new CpnFiles();
        $tree     = $CpnFiles->model_tree($mainTask['model_id'], $main_task_id);
        $trees    = $tree['model_tree'];
        $data     = !empty($trees) ? current($trees) : [];

        return $this->success($data);
    }

    /**
     * 双清单 - 数据列表
     * @param Request $request
     * @return mixed
     */
    public function getCpnData(Request $request)
    {
        $main_task_id            = $request->get('main_task_id', null);
        $cpn_type                = $request->get('cpn_type', null);//国产、进口
        $structure_id            = $request->get('structure_id', null); //型号节点
        $category                = $request->get('category', null); //当前元器件分类
        $cpn_specification_model = $request->get('cpn_specification_model', null); //型号规格
        $page                    = $request->get('page', 1); //页码
        $pagesize                = $request->get('pagesize', 10); //分页数据数量

        if (empty($main_task_id) || empty($cpn_type))
            return $this->error('参数不能为空');

        if (!$mainTask = MainTask::find($main_task_id))
            $this->error('信息不存在');

        $Cpn   = ($cpn_type == 2) ? new CpnImport() : new CpnDomestic();//进口数据表 / 国产数据表
        $where = ' 1 = 1 ';
        if (!empty($category))
        {
            $sqlType = env('DB_TYPE', 'mysql');
            //不同数据库的兼容性
            if ($sqlType === 'sqlite')
            {
                $where .= " and  substr(cpn_category_code, 0, 5) = {$category} ";
            } else
            {
                $where .= " and  LEFT(cpn_category_code, 4) = {$category} ";
            }
        }

        if (!empty($cpn_specification_model))
        {
            $where .= " and cpn_specification_model like  '%{$cpn_specification_model}%' ";
        }

        if (!empty($structure_id))
        {
            //选择某个节点
            $ModelStructure = new ModelStructure();
            if ($lists_str = $ModelStructure->getListsStr($structure_id, $main_task_id))
            {
                $where .= " and list_id in ({$lists_str}) ";
            } else
            {
                return $this->success(1);
            }
        } else
        {
            //页面初始化时
            if ($listIds = CpnFiles::where('main_task_id', $main_task_id)->pluck('id')->toArray())
            {
                $lists_str = implode(',', $listIds);
                $where     .= " and list_id in ({$lists_str}) ";
            } else
            {
                return $this->success(2);
            }
        }


        if ($cpn_type == 1)
        {
            $where .= " and is_repeat <> 1 ";
            $where .= " and result_grc <> 1 ";

            $select1 = [
                'cpn_import.list_id',
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
                'cpn_import.equip_use_number',
                'pro_review_opinion_import.pro_conclusion as result',
                'cpn_import.remark',
            ];

            $select2 = [
                'list_id',
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
                'yield_is_core_important as cpn_is_core_important',
                'safe_color',
                'equip_use_number',
                'result_grc as result',
                'remark',
            ];

            $listArr = explode(',',$lists_str);

            $addQuery = ProReviewOpinionAdd::leftJoin('pro_review_opinion_import', 'pro_review_opinion_add.hash_code', '=', 'pro_review_opinion_import.hash_code')
                ->leftJoin('cpn_import',function($join) use($listArr){
                    $join->on('pro_review_opinion_import.cpn_id', '=', 'cpn_import.id');
                    $join->whereIn('pro_review_opinion_import.list_id',$listArr);
                })
                ->where(['pro_review_opinion_add.main_task_id' => $main_task_id, 'pro_review_opinion_add.type' => 1])
                ->select($select1);

            $cpnList = $Cpn::whereRaw($where)->select($select2)->union($addQuery)->paginate($pagesize)->toArray();

            foreach ($cpnList['data'] as $key => $value)
            {
                if ($value['result'] == '国产替代纳入国产清单')
                {
                    $cpnList['data'][$key]['result'] = 4;
                }
                else
                {
                    return $this->error('错误');
                }
            }
        }
        elseif ($cpn_type == 2)
        {
            $select2 = [
                'id',
                'cpn_category_code',
                'cpn_manufacturer',
                'cpn_specification_model',
                'cpn_name',
                'cpn_quality',
                'cpn_package',
                'cpn_ref_price',
                'cpn_period',
                'cpn_country',
                'equip_use_number',
                'cpn_detect_apartment',
                'status as cpn_status',
                'cpn_control_level',
                'yield_is_core_important as cpn_is_core_important',
                'yield_safe_color as safe_color',
                'yield_proposed_safe_color as proposed_safe_color',
                'result_grc as result',
                'yield_necessity as necessity',
                'dependence',
                'remark',
            ];

            $where .= " and is_repeat <> 1 ";
            $where .= " and result_grc <> 1 ";
            $where .= " and result_unite in (2,3,4) ";

            $cpnList = $Cpn::whereRaw($where)->select($select2)->with(['ProReviewOpinionImport' => function ($query) {
                $query->where('pro_conclusion', '替换其他规格型号纳入进口清单');
                $query->with(['ProReviewOpinionAdd' => function ($addQuery) {
                    $addQuery->where('is_pass', 1);
                }]);
            }])->paginate($pagesize)->toArray();


            foreach ($cpnList['data'] as $cpnKey => $cpnValue)
            {
                if (!empty($cpnValue['pro_review_opinion_import']))
                {
                    if (!empty($cpnValue['pro_review_opinion_import']['pro_review_opinion_add']))
                    {
                        $addCpn                                              = $cpnValue['pro_review_opinion_import']['pro_review_opinion_add'];
                        $cpnList['data'][$cpnKey]['cpn_manufacturer']        = $addCpn['cpn_manufacturer_replace'];
                        $cpnList['data'][$cpnKey]['cpn_specification_model'] = $addCpn['cpn_specification_model_replace'];
                        $cpnList['data'][$cpnKey]['cpn_category_code']       = $addCpn['cpn_category_code'];
                        $cpnList['data'][$cpnKey]['cpn_name']                = $addCpn['cpn_name'];
                        $cpnList['data'][$cpnKey]['cpn_quality']             = $addCpn['cpn_quality'];
                        $cpnList['data'][$cpnKey]['cpn_package']             = $addCpn['cpn_package'];
                        $cpnList['data'][$cpnKey]['cpn_control_level']       = $addCpn['cpn_control_level'];
                        $cpnList['data'][$cpnKey]['cpn_ref_price']           = $addCpn['cpn_ref_price'];
                        $cpnList['data'][$cpnKey]['cpn_period']              = $addCpn['cpn_period'];
                        $cpnList['data'][$cpnKey]['cpn_detect_apartment']    = $addCpn['cpn_detect_apartment'];
                        $cpnList['data'][$cpnKey]['cpn_status']              = $addCpn['cpn_status'];
                        $cpnList['data'][$cpnKey]['cpn_country']             = $addCpn['cpn_country'];
                        $cpnList['data'][$cpnKey]['safe_color']              = $addCpn['safe_color'];
                        $cpnList['data'][$cpnKey]['proposed_safe_color']     = $addCpn['proposed_safe_color'];
                        $cpnList['data'][$cpnKey]['access_channel']          = $addCpn['access_channel'];
                        $cpnList['data'][$cpnKey]['necessity']               = $addCpn['necessity'];
                    }
                }

                unset($cpnList['data'][$cpnKey]['pro_review_opinion_import']);
            }
        }

        //审查规则x
        $ruleIds                          = AuxiliaryConfig::where('main_task_id', $main_task_id)->pluck('rule_id')->toArray();
        $cpnList['is_show_cross_version'] = in_array(1, $ruleIds) ? true : false;
        $cpnList['is_show_ycd']           = in_array(7, $ruleIds) ? true : false;
        $cpnList['is_show_cots']          = in_array(8, $ruleIds) ? true : false;

        return $this->success($cpnList);
    }

    /**
     * 生成数据包
     * @param Request $request
     * @return \Illuminate\Http\JsonResponse
     */
    public function makePackage(Request $request)
    {
        $main_task_id = $request->get('main_task_id', null);//主任务id

        if (!$task = MainTask::find($main_task_id))
            return $this->error('数据不存在');

        try
        {
            $model  = new ModelStructure();
            $result = $model->generatePackage($main_task_id);
            if ($result['success'] === true)
            {
                return $this->success($result['download_path']);
            } else
            {
                return $this->error($result['error_msg']);
            }
        } catch (\Exception $e)
        {
            return $this->error($e->getMessage() . $e->getTraceAsString());
        }
    }

    public function asd(Request $request){
        $file = $request->file('file');
        $fileName = $file->getClientOriginalName();
        $rootPath     = public_path('attachment/unPackage');

        $file->move($rootPath,$fileName);

        $zipName = $rootPath.DIRECTORY_SEPARATOR.$fileName;
        $toName = $rootPath.DIRECTORY_SEPARATOR.substr($fileName,0,strlen($fileName)-4);

        
        $zip = new \ZipArchive();
        $zip->open($zipName);
        $zip->extractTo($toName);
        $zip->close($toName);
    }
}