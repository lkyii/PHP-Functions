<?php

namespace App\Model;

use Illuminate\Database\Eloquent\Model;

class CpnList extends Model
{
    protected $table = 'cpn_list';

    public $cpnList = [
        'no' => ['序号'],
        'cpn_name' => ['元器件名称'],
        'cpn_specification_model' => ['规格型号'],
        'cpn_manufacturer' => ['生产厂商'],
        'cpn_quality' => ['质量等级'],
        'cpn_package' => ['封装形式'],
        'cpn_type' => ['国产 / 进口'],
        'temp_range' => ['温度范围'],
        'main_param' => ['主要性能参数'],
        'task_env' => ['任务环境'],
        'results' => ['试验验证结果'],
        'history' => ['应用经历'],
    ];

     /**
     * 关联参数
     * @return \Illuminate\Database\Eloquent\Relations\HasMany
     */
    public function params()
    {
        return $this->hasMany('App\Model\CpnParams', 'cpn_id', 'id');
    }
}
