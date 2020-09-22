<?php

namespace App\Http\Controllers;

use App\Model\CpnList;
use App\Model\Export;
use Illuminate\Http\JsonResponse;
use Illuminate\Http\Request;

class ExportController extends Controller
{
    /**
     * 导出元器件excel
     * @param Request $request
     * @return JsonResponse
     */
    public function export(Request $request)
    {
        $ids = $request->input('ids', null);

        $ids = explode(',', $ids);

        if (empty($ids))
            return $this->error('参数不能为空');

        $cpnList = CpnList::whereIn('id', $ids)->with(['params'])->get()->toArray();

        foreach ($cpnList as $key => $value) {
            $mainParam = '';

            foreach ($value['params'] as $mainKey => $mainValue) {
                $mainParam .= $mainValue['param_name'] . ':' . $mainValue['param_value'] . ';';
            }

            $cpnList[$key]['main_param'] = $mainParam;
        }

        $excelDir = 'attachment/export';
        $fileName = '元器件信息.xls';

        if (!file_exists($excelDir)) {
            $makeResult = mkdir($excelDir, 0777, true);

            if (!$makeResult)
                return $this->error('err 001');
        }

        $export = new Export();
        $result = $export->exportExcel($cpnList, $excelDir, $fileName);

        if ($result['status'] === true)
            return $this->success($result, '获取成功');

        return $this->error('导出失败', $result);
    }
}
