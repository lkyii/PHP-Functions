<?php

namespace App\Model;

use Illuminate\Support\Collection;
use PHPExcel_Cell;
use PHPExcel_Reader_Excel5;
use PHPExcel_RichText;

class Excel
{
    public $Title = [
        'no'                      => ['序号', '编号'],
        'cpn_category_code'       => ['分类代码', '分类', 'cate', 'category'],
        'cpn_name'                => ['元器件名称', '名称', 'name'],
        'cpn_specification_model' => ['规格型号', '型号规格', '型号', 'model', 'specification_model'],
        'cpn_manufacturer'        => ['生产厂商', 'manufacturer', '厂商', '生产厂家', '厂家', '建议生产商'],
        'cpn_quality'             => ['质量等级', 'quality', '质量'],
        'cpn_package'             => ['封装形式', '封装', 'footprint', 'package'],
        'cpn_control_level'       => ['自主可控等级', '可控等级', 'level'],
        'cpn_is_core_important'   => ['是否核心关键器件', '核心关键器件', '关键器件', 'core'],
        'cpn_ref_price'           => ['参考价格', '价格', 'price'],
        'cpn_period'              => ['供货周期', '周期', 'period'],
        'cpn_detect_apartment'    => ['元器件检测机构', '检测机构', 'apartment'],
        'history'                 => ['历史应用信息', '应用信息'],
        'equip_use_number'        => ['分系统单设备使用量', '单设备使用量'],
        //        'equip_number'                   => ['分系统设备台套数', '设备台套数'],
        //        'equip_name'                     => ['分系统设备名称', '设备名称'],
        //        'equip_research_apartment'       => ['分系统设备研制单位', '设备研制单位'],
        'remark'                  => ['备注', 'remark'],
        'cpn_country'             => ['国别地区', '国家地区', 'country'],
        'safe_color'              => ['安全颜色等级', 'anquan颜色等级', 'anquan颜色', 'anquan等级颜色'],
        'proposed_safe_color'     => ['建议安全等级颜色', '建议anquan等级颜色', 'jianyianquan颜色', 'jianyianquan等级颜色'],
        'necessity'               => ['必要性'],
        'access_channel'          => ['获取渠道', '渠道'],
    ];

    // 国产表头转换
    public $domesticTitle = [
        'no'                       => ['序号', '编号'],
        'cpn_category_code'        => ['分类代码', '分类', 'cate', 'category'],
        'cpn_name'                 => ['元器件名称', '名称', 'name'],
        'cpn_specification_model'  => ['规格型号', '型号规格', '型号', 'model', 'specification_model'],
        'equip_use_number'         => ['单机使用数量', '单设备使用量'],
        'cpn_manufacturer'         => ['生产厂商', 'manufacturer', '厂商', '生产厂家', '厂家', '建议生产商'],
        'cpn_quality'              => ['质量等级', 'quality', '质量'],
        'cpn_package'              => ['封装形式', '封装', 'footprint', 'package'],
        'cpn_control_level'        => ['自主可控等级', '可控等级', 'level'],
        'cpn_is_core_important'    => ['是否核心关键器件', '核心关键器件', '关键器件', 'core'],
        'cpn_ref_price'            => ['参考价格', '价格', 'price'],
        'cpn_period'               => ['供货周期', '周期', 'period'],
        'cpn_detect_apartment'     => ['元器件检测机构', '检测机构', 'apartment'],
        'history'                  => ['历史应用信息', '应用信息'],
        //        'equip_number'             => ['分系统设备台套数', '设备台套数'],
        //        'equip_name'               => ['分系统设备名称', '设备名称'],
        //        'equip_research_apartment' => ['分系统设备研制单位', '设备研制单位'],
        'remark'                   => ['备注', 'remark'],
        'equip_name'               => ['分系统/设备名称', 'equip_name'],
        'equip_research_apartment' => ['分系统/设备研制单位', 'equip_research_apartment'],

        // 新增专家意见列
        'kb_aux_satisfaction_msg'  => ['历史满足度审查情况', 'kb_aux_satisfaction_msg'],
        'kb_satisfaction_whether'  => ['是否存在风险（历史质量问题风险、研制生产方面的固有缺陷和风险、元器件应用常见电路风险、个性化风险，如防静电能力低、专题化风险，如抗辐射能力等）', 'kb_satisfaction_whether'],
        'kb_staisfaction_msg'      => ['风险概述', 'kb_staisfaction_msg'],
        'kb_aux_empty_whether'     => ['是否存在伪空包现象', 'kb_aux_empty_whether'],
        'kb_aux_empty_msg'         => ['历史伪空包现象审查情况', 'kb_aux_empty_msg'],
        'kb_empty_whether'         => ['是否存在伪空包问题', 'kb_empty_whether'],
        'kb_empty_msg'             => ['问题概述', 'kb_empty_msg'],
        'hash_code'                => ['哈希值', 'hash_code'],
    ];

    // 意见闭环 - 国产表头转换
    public $domesticTitle2 = [
        'no'                          => ['序号', '编号'],
        'cpn_category_code'           => ['分类代码', '分类', 'cate', 'category'],
        'cpn_name'                    => ['元器件名称', '名称', 'name'],
        'cpn_specification_model'     => ['规格型号', '型号规格', '型号', 'model', 'specification_model'],
        'equip_use_number'            => ['单机使用数量', '单设备使用量'],
        'cpn_manufacturer'            => ['生产厂商', 'manufacturer', '厂商', '生产厂家', '厂家', '建议生产商'],
        'cpn_quality'                 => ['质量等级', 'quality', '质量'],
        'cpn_package'                 => ['封装形式', '封装', 'footprint', 'package'],
        'cpn_control_level'           => ['自主可控等级', '可控等级', 'level'],
        'cpn_is_core_important'       => ['是否核心关键器件', '核心关键器件', '关键器件', 'core'],
        'cpn_ref_price'               => ['参考价格', '价格', 'price'],
        'cpn_period'                  => ['供货周期', '周期', 'period'],
        'cpn_detect_apartment'        => ['元器件检测机构', '检测机构', 'apartment'],
        'history'                     => ['历史应用信息', '应用信息'],
        'remark'                      => ['备注', 'remark'],
        'equip_name'                  => ['分系统/设备名称', 'equip_name'],
        'equip_research_apartment'    => ['分系统/设备研制单位', 'equip_research_apartment'],
        'kb_satisfaction_whether'     => ['是否存在风险', 'kb_satisfaction_whether'],
        'kb_satisfaction_expert_msg'  => ['满足度审查意见', 'kb_satisfaction_expert_msg'],
        'satisfaction_avoid_measure'  => ['满足度风险规避措施', 'satisfaction_avoid_measure'],
        'kb_empty_whether'            => ['是否存在伪空包问题', 'kb_empty_whether'],
        'kb_empty_expert_msg'         => ['伪空包现象审查情况', 'kb_empty_expert_msg'],
        'is_agree_empty_package'      => ['是否采纳', 'is_agree_empty_package'],
        'empty_package_avoid_measure' => ['伪空包现象规避措施', 'empty_package_avoid_measure'],
        'hash_code'                   => ['哈希值']
    ];

    // 国产校验规则
    public $verifyDomesticRules = [
        'header'     => [
            "no"                      => "A",
            "cpn_category_code"       => "B",
            "cpn_name"                => "C",
            "cpn_specification_model" => "D",
            "equip_use_number"        => "E",
            "cpn_manufacturer"        => "F",
            "cpn_quality"             => "G",
            "cpn_package"             => "H",
            "cpn_control_level"       => "I",
            "cpn_is_core_important"   => "J",
            "cpn_ref_price"           => "K",
            "cpn_period"              => "L",
            "cpn_detect_apartment"    => "M",
            "history"                 => "N",
            "remark"                  => "O",
        ],
        'required'   => [
            //            "no",
            //            "cpn_category_code",
            "cpn_name",
            //            "cpn_specification_model",
            //            "cpn_manufacturer",
            "cpn_quality",
            "cpn_package",
            "cpn_control_level",
            "cpn_is_core_important",
            "cpn_ref_price",
            "cpn_period",
            "cpn_detect_apartment",
            "equip_use_number",
        ],
        'nopass'     => [
            "cpn_category_code",
            "cpn_specification_model",
            "cpn_manufacturer",
        ],
        'dataType'   => [
            'fields' => [
                "no",
                "cpn_category_code",
                //                "cpn_ref_price",
                //                "cpn_period",
                "equip_use_number",
            ],
            'rules'  => [
                "no"                => "int",
                "cpn_category_code" => "int",
                //                "cpn_ref_price"     => "float",
                //                "cpn_period"        => "float",
                "equip_use_number"  => "int",
            ]

        ],
        'dataFormat' => [ //数据格式
            "cpn_ref_price",
            "cpn_period",
        ],
        'scope'      => [
            'fields'   => [
                "cpn_is_core_important",
                "cpn_control_level",
                "cpn_detect_apartment",
            ],
            'category' => [
                "cpn_category_code",
            ],
            'rules'    => [
                "cpn_is_core_important" => [0, 1],
                "cpn_control_level"     => ['A', 'B', 'C', 'D', 'E'],
                "cpn_detect_apartment"  => ['军用电子元器件第一检测中心', '军用电子元器件第二检测中心', '军用电子元器件第三检测中心', '广州检测中心', '军用电子元器件检测技术研究中心', '中国运载火箭研究院物流中心', '中国空间技术研究院宇航物资保障事业部', '中国科学院空间应用工程与技术中心', '其它'],
            ]
        ]
    ];

    // 进口表头转换
    public $importTitle = [
        'no'                         => ['序号', '编号'],
        'equip_name'                 => ['装机信息', '设备名称'],
        'cpn_category_code'          => ['元器件类别', '分类', 'cate', 'category'],
        'cpn_name'                   => ['元器件名称', '名称', 'name'],
        'cpn_specification_model'    => ['规格型号', '型号规格', '型号', 'model', 'specification_model'],
        //        'equip_use_number'           => ['单机使用数量', '单设备使用量'],
        'cpn_manufacturer'           => ['生产厂商', 'manufacturer', '厂商', '生产厂家', '厂家', '建议生产商'],
        'cpn_country'                => ['国别地区', '国家地区', 'country'],
        'cpn_quality'                => ['质量等级', 'quality', '质量'],
        'cpn_package'                => ['封装形式', '封装', 'footprint', 'package'],
        'safe_color'                 => ['安全颜色等级', 'anquan颜色等级', 'anquan颜色', 'anquan等级颜色'],
        'proposed_safe_color'        => ['建议安全等级颜色', '建议anquan等级颜色', 'jianyianquan颜色', 'jianyianquan等级颜色'],
        'cpn_is_core_important'      => ['是否核心关键器件', '核心关键器件', '关键器件', 'core'],
        'necessity'                  => ['必要性'],
        'result_pc'                  => ['计算机辅助比对审查意见'],
        //        'cpn_ref_price'              => ['参考价格', '价格', 'price'],
        //        'cpn_period'                 => ['供货周期', '周期', 'period'],
        //        'access_channel'             => ['获取渠道', '渠道'],
        //        'equip_number'               => ['分系统设备台套数', '设备台套数'],
        //        'equip_name'                 => ['分系统设备名称', '设备名称'],
        //        'equip_research_apartment'   => ['分系统设备研制单位', '设备研制单位'],
        //        'remark'                     => ['备注', 'remark'],
        //        'equip_name'                 => ['分系统/设备名称', 'equip_name'],
        //        'equip_research_apartment'   => ['分系统/设备研制单位', 'equip_research_apartment'],
        //        'kb_mast_whether'            => ['是否认可研制单位选用必要性说明', 'proKnowlege'],
        //        'kb_mast_msg'                => ['判定必要性不足的理由(选择否时填写)', 'proKnowlege'],
        //        'kb_aux_msg'                 => ['历史安全性审查情况', 'proKnowlege'],
        //        'kb_safety_whether'          => ['是否存在安全风险（确定有安全隐患或可能存在安全隐患）', 'proKnowlege'],
        //        'kb_safety_color'            => ['建议安全等级颜色', 'proKnowlege'],
        //        'kb_safety_msg'              => ['建议安全等级颜色的理由', 'proKnowlege'],
        //        'kb_aux_insurability_msg'    => ['历史可保障性审查情况', 'proKnowlege'],
        //        'kb_insurability_whether'    => ['是否存在供货风险', 'proKnowlege'],
        //        'kb_insurability_color'      => ['建议安全等级颜色', 'proKnowlege'],
        //        'kb_insurability_msg'        => ['建议安全等级颜色的理由', 'proKnowlege'],
        //        'kb_aux_substitution_plan'   => ['替代方式', 'proKnowlege'],
        'kb_aux_substitution_model'  => ['可替代产品型号(厂家/质量等级）'],
        'kb_aux_substitution_status' => ['可替代产品应用状态'],
        //        'kb_aux_substitution_msg'    => ['历史可替代产品审查情况''],
        'kb_substitution_plan'       => ['替代方案选择（选择一种最优方案）'],
        'kb_substitution_model'      => ['可替代产品型号（选择认可国产化替代报告方案或补充其他时填写）'],
        'kb_substitution_mfr'        => ['可替代产品厂商（选择认可国产化替代报告方案或补充其他时填写）'],
        'kb_substitution_whether'    => ['替代类型', 'proKnowlege'],
        'kb_substitution_status'     => ['可替代产品状态', 'proKnowlege'],
        'pro_massage'                => ['专家审查结果', 'proKnowlege'],
        'pro_massage_status'         => ['是否接受专家（或计算机比对）审查意见', 'proKnowlege'],
        'pro_way'                    => ['具体处理措施', 'proKnowlege'],
        'pro_conclusion'             => ['最终审查结论', 'proKnowlege'],
        'hash_code'                  => ['哈希值', 'hash_code'],
    ];

    public $importTitleExt01 = [
        'hash_code'                       => ['用研审查清单序号', '分类', 'cate', 'category'],
        'cpn_specification_model'         => ['型号规格', '分类', 'cate', 'category'],
        'cpn_category_code'               => ['分类代码', '分类', 'cate', 'category'],
        'cpn_name'                        => ['元器件名称', '名称', 'name'],
        'cpn_specification_model_replace' => ['替代规格型号', '型号规格', '型号', 'model', 'specification_model'],
        'cpn_manufacturer_replace'        => ['替代生产厂商', 'manufacturer', '厂商', '生产厂家', '厂家', '建议生产商'],
        'cpn_quality'                     => ['质量等级', 'quality', '质量'],
        'cpn_package'                     => ['封装形式', '封装', 'footprint', 'package'],
        'cpn_control_level'               => ['自主可控等级'],
        'cpn_ref_price'                   => ['参考价格', '价格', 'price'],
        'cpn_period'                      => ['供货周期', '周期', 'period'],
        'cpn_detect_apartment'            => ['元器件检测机构', '周期', 'period'],
        'cpn_replace_status'              => ['替代类型', '周期', 'period'],
        'cpn_status'                      => ['产品状态', '周期', 'period'],
    ];

    public $importTitleExt02 = [
        'hash_code'                       => ['用研审查清单序号', '分类', 'cate', 'category'],
        'cpn_specification_model'         => ['型号规格', '分类', 'cate', 'category'],
        'cpn_category_code'               => ['分类代码', '分类', 'cate', 'category'],
        'cpn_name'                        => ['元器件名称', '名称', 'name'],
        'cpn_specification_model_replace' => ['替代规格型号', '型号规格', '型号', 'model', 'specification_model'],
        'cpn_manufacturer_replace'        => ['替代生产厂商', 'manufacturer', '厂商', '生产厂家', '厂家', '建议生产商'],
        'cpn_country'                     => ['国别地区', 'manufacturer', '厂商', '生产厂家', '厂家', '建议生产商'],
        'cpn_quality'                     => ['质量等级', 'quality', '质量'],
        'cpn_package'                     => ['封装形式', '封装', 'footprint', 'package'],
        'safe_color'                      => ['安全等级颜色'],
        'proposed_safe_color'             => ['建议安全等级颜色'],
        'necessity'                       => ['必要性'],
        'cpn_ref_price'                   => ['参考价格', '价格', 'price'],
        'cpn_period'                      => ['供货周期', '周期', 'period'],
        'access_channel'                  => ['获取渠道'],
    ];

    // 意见闭环 - 进口表头转换
    public $importTitle2 = [
        'no'                       => ['序号', '编号'],
        'cpn_category_code'        => ['分类代码', '分类', 'cate', 'category'],
        'cpn_name'                 => ['元器件名称', '名称', 'name'],
        'cpn_specification_model'  => ['规格型号', '型号规格', '型号', 'model', 'specification_model'],
        'equip_use_number'         => ['单机使用数量', 'equip_use_number'],
        'cpn_manufacturer'         => ['生产厂商', 'manufacturer', '厂商', '生产厂家', '厂家', '建议生产商'],
        'cpn_country'              => ['国别地区', '国家地区', 'country'],
        'cpn_quality'              => ['质量等级', 'quality', '质量'],
        'cpn_package'              => ['封装形式', '封装', 'footprint', 'package'],
        'safe_color'               => ['安全颜色等级', 'anquan颜色等级', 'anquan颜色', 'anquan等级颜色'],
        'proposed_safe_color'      => ['建议安全等级颜色', '建议anquan等级颜色', 'jianyianquan颜色', 'jianyianquan等级颜色'],
        'cpn_is_core_important'    => ['是否核心关键器件', '核心关键器件', '关键器件', 'core'],
        'necessity'                => ['必要性'],
        'cpn_ref_price'            => ['参考价格', '价格', 'price'],
        'cpn_period'               => ['供货周期', '周期', 'period'],
        'access_channel'           => ['获取渠道', '渠道'],
        'remark'                   => ['备注', 'remark'],
        'equip_name'               => ['分系统设备名称', '设备名称'],
        'equip_research_apartment' => ['分系统设备研制单位', '设备研制单位'],

        'kb_mast_whether'              => ['是否认可研制单位选用必要性说明', 'kb_mast_whether'],
        'kb_mast_expert_msg'           => ['必要性审查意见', 'kb_mast_expert_msg'],
        'is_continue_choose'           => ['是否继续选用', 'is_continue_choose'],
        'continue_choose_reason'       => ['替换选用措施或继续选用必要性补充说明', 'continue_choose_reason'],
        'kb_safety_whether'            => ['是否存在安全风险', 'kb_safety_whether'],
        'kb_safety_expert_msg'         => ['安全性审查意见', 'kb_safety_expert_msg'],
        'safety_color_grade'           => ['安全性确认安全等级颜色', 'safety_color_grade'],
        'satefy_avoid_measure'         => ['安全性风险规避措施', 'satefy_avoid_measure'],
        'kb_insurability_whether'      => ['是否存在供货风险', 'kb_insurability_whether'],
        'kb_insurability_expert_msg'   => ['可保障性审查意见', 'kb_insurability_expert_msg'],
        'insurability_color_grade'     => ['可保障性确认安全等级颜色', 'insurability_color_grade'],
        'insurability_avoid_measure'   => ['可保障性风险规避措施', 'insurability_avoid_measure'],
        'kb_substitution_expert_msg'   => ['国产化替代审查专家意见', 'kb_substitution_expert_msg'],
        //        'kb_substitution_plan'         => ['国产替代化审查意见(替代方案选择)', 'kb_substitution_plan'],
        //        'kb_substitution_whether'      => ['国产化替代审查专家意见', 'kb_substitution_whether'],
        //        'kb_substitution_status'       => ['可替代产品状态', 'kb_substitution_status'],
        'is_agree_substitution'        => ['是否同意国产化替代意见', 'is_agree_substitution'],
        'choose_replace_spe'           => ['选择可替代型号', 'choose_replace_spe'],
        'replace_jujde'                => ['替代性判别', 'replace_jujde'],
        'no_agree_substitution_reason' => ['不采纳国产替代性意见理由说明', 'no_agree_substitution_reason'],
        'hash_code'                    => ['哈希值', 'hash_code']
    ];

    // 进口校验规则
    public $verifyImportRules = [
        'header'     => [
            "no"                      => "A",
            "cpn_category_code"       => "B",
            "cpn_name"                => "C",
            "cpn_specification_model" => "D",
            "equip_use_number"        => "E",
            "cpn_manufacturer"        => "F",
            "cpn_country"             => "G",
            "cpn_quality"             => "H",
            "cpn_package"             => "I",
            "safe_color"              => "J",
            "proposed_safe_color"     => "K",
            "cpn_is_core_important"   => "L",
            "necessity"               => "M",
            "cpn_ref_price"           => "N",
            "cpn_period"              => "O",
            "access_channel"          => "P",
            'remark'                  => "Q",
        ],
        'required'   => [
            "no",
            //            "cpn_category_code",
            "cpn_name",
            //            "cpn_specification_model",
            //            "cpn_manufacturer",
            "cpn_country",
            "cpn_quality",
            "cpn_package",
            "safe_color",
            "proposed_safe_color",
            "cpn_is_core_important",
            "necessity",
            "cpn_ref_price",
            "cpn_period",
            "equip_use_number",
        ],
        'nopass'     => [
            "cpn_category_code",
            "cpn_specification_model",
            "cpn_manufacturer",
        ],
        'dataType'   => [
            'fields' => [
                "no",
                "cpn_category_code",
                //                "cpn_ref_price",
                //                "cpn_period",
                "equip_use_number",
            ],
            'rules'  => [
                "no"                => "int",
                "cpn_category_code" => "int",
                //                "cpn_ref_price"     => "float",
                //                "cpn_period"        => "float",
                "equip_use_number"  => "int",
            ]
        ],
        'dataFormat' => [ //数据格式
            "cpn_ref_price",
            "cpn_period",
        ],
        'scope'      => [
            'fields'   => [
                'cpn_is_core_important',
                'necessity',
                'access_channel',
                'safe_color',
                'proposed_safe_color',
            ],
            'category' => ['cpn_category_code'],
            'rules'    => [
                'cpn_is_core_important' => [0, 1],
                'necessity'             => ['1', '2.1', '2.2', '2.3', '3', '4'],
                'access_channel'        => ['直接', '第三方', '秘密'],
                'safe_color'            => ['红色', '紫色', '橙色', '黄色', '绿色', '红', '紫', '橙', '黄', '绿'],
                'proposed_safe_color'   => ['红色', '紫色', '橙色', '黄色', '绿色', '红', '紫', '橙', '黄', '绿'],
            ],

        ]
    ];

    // 合规审查问题清单 - 国产表头
    public $domesticTitle3 = [
        'no'                          => "序号",
        'zhuang_position'             => "装机信息",
        'cpn_category_code'           => '分类代码',
        'cpn_name'                    => '元器件名称',
        'cpn_specification_model'     => '规格型号',
        'equip_use_number'            => '单机使用数量',
        'cpn_manufacturer'            => '生产厂商',
        'cpn_quality'                 => '质量等级',
        'cpn_package'                 => '封装形式',
        'cpn_control_level'           => '自主可控等级',
        'cpn_is_core_important'       => '是否核心关键器件',
        'cpn_ref_price'               => '参考价格',
        'cpn_period'                  => '供货周期',
        'cpn_detect_apartment'        => '元器件检测机构',
        'history'                     => '历史应用信息',
        'remark'                      => '备注',
    ];

    // 合规审查问题清单 - 进口表头
    public $importTitle3 = [
        'no'                       => '序号',
        'zhuang_position'          => "装机信息",
        'cpn_category_code'        => '分类代码',
        'cpn_name'                 => '元器件名称',
        'cpn_specification_model'  => '规格型号',
        'equip_use_number'         => '单机使用数量',
        'cpn_manufacturer'         => '生产厂商',
        'cpn_country'              => '国别地区',
        'cpn_quality'              => '质量等级',
        'cpn_package'              => '封装形式',
        'safe_color'               => '安全颜色等级',
        'proposed_safe_color'      => '建议安全等级颜色',
        'cpn_is_core_important'    => '是否核心关键器件',
        'necessity'                => '必要性',
        'cpn_ref_price'            => '参考价格',
        'cpn_period'               => '供货周期',
        'access_channel'           => '获取渠道',
        'remark'                   => '备注'
    ];



    //线缆类分类
    public $LineCategories = [
        '5995',
        '6010',
        '6020'
    ];

    public $CotsData = [
        '工业级',
        '商业级',
        '扩展工业级',
        '扩展商业级',
        '汽车工业级',
        '军温工业级'
    ];

    public $ControlData = [
        1 => 'A',
        2 => 'B',
        3 => 'C',
        4 => 'D',
        5 => 'E' //最低级
    ];

    public $ColorGrade = [
        1 => '绿色',
        2 => '黄色',
        3 => '橙色',
        4 => '紫色',
        5 => '红色', //最严
    ];



    /**
     * 数字转字母 （类似于Excel列标）
     * @param Int $index 索引值
     * @param Int $start 字母起始值
     * @return String 返回字母
     */
    function IntToChr($index, $start = 65)
    {
        $str = '';
        if (floor($index / 26) > 0)
        {
            $str .= $this->IntToChr(floor($index / 26) - 1);
        }
        return $str . chr($index % 26 + $start);
    }

    /**
     * 解析excel文件
     * @param $filePath
     * @return array
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     */
    public function readExcel($filePath, $readNum = 30, $readRowNum = 2, $sheetArr = [0])
    {

        // 建立reader对象
        $phpReader = new \PHPExcel_Reader_Excel2007();
        if (!$phpReader->canRead($filePath))
        {
            $phpReader = new PHPExcel_Reader_Excel5();
            if (!$phpReader->canRead($filePath))
            {
                return ['error_code' => 40001, 'msg' => 'Excel读取错误', 'data' => ''];
            }
        }

        // 建立excel对象
        $phpExcel = $phpReader->load($filePath);//读取文件

        // 总页数
        $sheetCount = $phpExcel->getSheetCount();

        $allSheetData = [];
        foreach ($sheetArr as $value)
        {
            $currentSheet = $phpExcel->getSheet($value);// 读取excel第一个工作表
            $currentTitle = $currentSheet->getTitle();//获取页脚

            // 获取最大列号
            $allColumn = $currentSheet->getHighestColumn();
            // 获取行数
            $allRow = $currentSheet->getHighestRow();

            $rowsData      = [];
            $colsData      = [];
            $title         = '';
            $header        = [];
            $headRow       = 1;
            $startRow      = 1;
            $emptys        = [
                'rows' => [],
                'cols' => [],
            ];
            $normalHeader  = []; // 正常表头
            $defaultHeader = [];
            $orderColStat  = 0; // 数字列

            $allColumnNum = PHPExcel_Cell::columnIndexFromString($allColumn);


            //循环读取每个单元格的内容。注意行从1开始，列从A开始
            for ($rowIndex = $readRowNum; $rowIndex <= $allRow; $rowIndex++)
            {
                // 行
                for ($colNum = 0; ($colNum < $allColumnNum) && $colNum < $readNum; $colNum++)
                {
                    // 列
                    $colIndex                 = PHPExcel_Cell::stringFromColumnIndex($colNum);//列 ABC.....
                    $defaultHeader[$colIndex] = $colIndex; // 表头
                    $addr                     = $colIndex . $rowIndex; // 结构:A1
                    $cell                     = $currentSheet->getCell($addr)->getCalculatedValue();//获取单元格的值
                    if ($cell instanceof PHPExcel_RichText)
                    {
                        // 富文本转字符串
                        $cell = $cell->__toString();
                    }
                    // 转化成UTF8
                    $cell = String2Utf8($cell);
                    // 全角转半角
                    $cell = trim(String2ASC($cell));
                    if (!strlen($cell) && (@$emptys['rows'][$rowIndex] || !isset($emptys['rows'][$rowIndex])))
                    {
                        $emptys['rows'][$rowIndex] = true;
                    } else
                    {
                        $emptys['rows'][$rowIndex] = false;
                    }

                    if (!strlen($cell) && (@$emptys['cols'][$colIndex] || !isset($emptys['cols'][$colIndex])))
                    {
                        $emptys['cols'][$colIndex] = true;
                    } else
                    {
                        $emptys['cols'][$colIndex] = false;
                    }

                    $rowsData[$rowIndex][$colIndex] = preg_replace("/^(\s|\&nbsp\;|　|\xc2\xa0)+/", "", $cell);//执行一个正则表达式的搜索和替换

                    $colsData[$colIndex][$rowIndex] = preg_replace("/^(\s|\&nbsp\;|　|\xc2\xa0)+/", "", $cell);
                    //第一列是否为序列
                    if ($colIndex == 'A' && (int)$cell > 1 && ($cell - 1) == $rowsData[$rowIndex - 1]['A'])
                    {//值为1时，转换int类型为0  获取序号一列
                        $orderColStat++;
                    }

                }
                // 判断是否为表头
                if ($rowIndex < 10 && $rowsData[$rowIndex] && !$normalHeader)
                {
                    $normalHeader = $this->_parseBomExcelHeader($rowsData[$rowIndex]);

                    // header 前面的行将不处理
                    if ($normalHeader)
                    {
                        $headRow  = $rowIndex;
                        $startRow = $rowIndex + 1;
                    }
                }
                if (!$header && $defaultHeader)
                {
                    $header = $defaultHeader;
                }
            }

            if (empty($emptys['rows']) || empty($emptys['cols']))
            {
                return ['error_code' => 50005, 'msg' => '上传数据为空或错误', 'data' => []];
            }

            // 删除空行
            foreach ($emptys['rows'] as $k => $v)
            {
                if ($v)
                {
                    unset($rowsData[$k]);
                }
            }

            // 删除空列
            foreach ($emptys['cols'] as $k => $v)
            {
                if ($v)
                {
                    unset($colsData[$k]);
                }
            }

            foreach ($colsData as $col => &$cValue)
            {
                foreach ($cValue as $k => &$v)
                {
                    if ($emptys['rows'][$k])
                    {
                        unset($cValue[$k]);
                    }
                }
            }

            // 判断第一列是否为序列，50%按顺序递增时
            if ($orderColStat / (count($rowsData)) > 0.5)
            {
                $orderCol           = 'A';
                $normalHeader['no'] = $orderCol;
            }

            $sheetData = [
                'title'             => $title,
                'currentSheetTitle' => $currentTitle,
                'sheetCount'        => $sheetCount,
                'header'            => $header,
                'headRow'           => $headRow,
                'baseHeader'        => $normalHeader,
                'maxCol'            => count($colsData),
                'maxRow'            => count($rowsData),
                'startRow'          => $startRow,
                'rowsData'          => $rowsData,
                'colsData'          => $colsData,
            ];

            array_push($allSheetData, $sheetData);
        }


        return ['error_code' => 0, 'msg' => '', 'data' => $allSheetData];
    }

    /**
     * 解析表头
     * @param $row excel行数据
     * @return array | | false 返回匹配header的对应关系
     */
    public function _parseBomExcelHeader($row)
    {
        if (!$row || count($row) < 2)
        {
            return [];
        }

        //序号，型号规格，数量，....
        $normalTemp   = $this->Title;
        $normalHeader = [];
        foreach ($row as $cellNo => $item)
        {

            //匹配关键字
            foreach ($normalTemp as $key => $words)
            {
                //已匹配
                if (isset($normalHeader[$key]) && $normalHeader[$key])
                {//如果变量已经存在  就结束
                    continue;
                }
                $item = preg_replace('/\(.*\)/', '', $item);
                $item = strtolower(str_replace([' ', '*', '-', '/'], '', $item));
                if (in_array($item, $words))
                {
                    $normalHeader[$key] = $cellNo;
                }
            }
        }

        return $normalHeader;
    }

    /**
     * 校验数据
     * @param $fileType
     * @param $data
     * @param $main_task_id
     * @return array
     */
    public function verifyData($fileType, $data, $main_task_id)
    {
        $errInfo      = [];
        $batchData    = [];
        $batchImpData = [];

        $category  = Cpncategories::pluck('id')->toArray();//分类
        $chunkSize = 1000;
        $cpn_type  = ($fileType == '国产') ? 1 : 2;
        if ($fileType == '国产')
        {
            $domesticRules  = $this->verifyDomesticRules;
            $domesticTitle3 = $this->domesticTitle3;
            //分批审查
            $chunks = collect($data)->chunk($chunkSize);
            $j      = 1;//记录错误行数
            foreach ($chunks as $block)
            {
                $rows = $block->toArray();
                $i    = 0;
                foreach ($rows as $k => $row)
                {
                    //初始化数组
                    $batchData[$row['id']] = [
                        'id'                      => $row['id'],
                        "yield_is_core_important" => @$row['cpn_is_core_important'],
                        "yield_control_level"     => @$row['cpn_control_level'],
                        'is_repeat'               => 0,
                        "result_grc"              => 3
                    ];
                    foreach ($row as $field => $val)
                    {

                        //判断关键信息必填项 是否为空
                        if (in_array($field, $domesticRules['nopass']) && (empty($val) && $val !== 0))
                        {
                            //问题行数递增
                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                            {
                                $j++;
                            }

                            $batchData[$row['id']]['result_grc'] = 1;

                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $row['id'],
                                'review_type'  => 1,
                                'messages'     => $domesticTitle3[$field].'信息缺失',
                                'row'          => $j,
                                'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                            ];

                            $i++;
                            break;

                        }


                        // 判断分类代码
                        if (in_array($field, $domesticRules['scope']['category']) && (!empty($val) || ($val === "0")) && !in_array($val, $category))
                        {
                            //问题行数递增
                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                            {
                                $j++;
                            }

                            //更新该条元器件审查状态
                            $batchData[$row['id']]["result_grc"] = 1;


                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $row['id'],
                                'review_type'  => 1,
                                'messages'     => '分类代码不存在',
                                'row'          => $j,
                                'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                            ];
                            $i++;
                            break;

                        }


                        // 判断 其他必填信息 是否为空
                        if (in_array($field, $domesticRules['required']) && (empty($val) && $val !== 0))
                        {
                            //问题行数递增
                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                            {
                                $j++;
                            }

                            //更新该条元器件审查状态
                            if ($field == "cpn_is_core_important")
                            {
                                $batchData[$row['id']]['yield_is_core_important'] = 1;
                                $batchData[$row['id']]['result_grc']              = 2;
                            }
                            if ($field == "cpn_control_level")
                            {
                                $batchData[$row['id']]['yield_control_level'] = "E";
                                $batchData[$row['id']]['result_grc']          = 2;
                            }

                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $row['id'],
                                'review_type'  => 2,
                                'messages'     => '其他必填信息不能为空',
                                'row'          => $j,
                                'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                            ];

                            $i++;
                            continue;

                        }

                        // 验证数据类型
                        if (in_array($field, $domesticRules['dataType']['fields']) && ($val != ''))
                        {
                            //转化成数值型
                            $numVal = $val;
                            if (is_string($numVal))
                            {
                                //字符串必须由数字和.组成
                                if (preg_match('/[^0-9\.]/', $numVal))
                                {
                                    //问题行数递增
                                    if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                    {
                                        $j++;
                                    }

                                    //更新该条元器件审查状态
                                    $batchData[$row['id']]['result_grc'] = 2;

                                    $errInfo[] = [
                                        'main_task_id' => $main_task_id,
                                        'list_id'      => $row['list_id'],
                                        'cpn_type'     => $cpn_type,
                                        'cpn_id'       => $row['id'],
                                        'review_type'  => 3,
                                        'messages'     => '数据类型不正确',
                                        'row'          => $j,
                                        'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                                    ];

                                    $i++;
                                    continue;
                                }
                                if (strpos($numVal, '.') !== false)
                                {
                                    $numVal = floatval($numVal);
                                } else
                                {
                                    $numVal = intval($numVal);
                                }
                            }

                            if ($numVal > 0)
                            {
                                if (($field !== 'cpn_ref_price') && !($field == 'equip_use_number' && in_array(substr($row['cpn_category_code'], 0, 4), $this->LineCategories)) && !is_int($numVal))//参考价格
                                {
                                    //问题行数递增
                                    if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                    {
                                        $j++;
                                    }

                                    //更新该条元器件审查状态
                                    $batchData[$row['id']]['result_grc'] = 2;

                                    $errInfo[] = [
                                        'main_task_id' => $main_task_id,
                                        'list_id'      => $row['list_id'],
                                        'cpn_type'     => $cpn_type,
                                        'cpn_id'       => $row['id'],
                                        'review_type'  => 3,
                                        'messages'     => '数据类型不正确',
                                        'row'          => $j,
                                        'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                                    ];

                                    $i++;
                                    continue;
                                }
                            } else
                            {
                                //问题行数递增
                                if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                {
                                    $j++;
                                }

                                //更新该条元器件审查状态
                                $batchData[$row['id']]['result_grc'] = 2;

                                $errInfo[] = [
                                    'main_task_id' => $main_task_id,
                                    'list_id'      => $row['list_id'],
                                    'cpn_type'     => $cpn_type,
                                    'cpn_id'       => $row['id'],
                                    'review_type'  => 3,
                                    'messages'     => '数据类型不正确',
                                    'row'          => $j,
                                    'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                                ];

                                $i++;
                                continue;
                            }

                        }

                        // 验证数据格式
                        if (in_array($field, $domesticRules['dataFormat']) && ($val != ''))
                        {
                            //转化成数值型
                            $numVal = $val;
                            if (is_string($numVal))
                            {
                                //字符串必须由数字和. _组成
                                if (preg_match('/[^0-9\._]/', $numVal))
                                {
                                    //问题行数递增
                                    if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                    {
                                        $j++;
                                    }

                                    //更新该条元器件审查状态
                                    $batchData[$row['id']]['result_grc'] = 2;

                                    $errInfo[] = [
                                        'main_task_id' => $main_task_id,
                                        'list_id'      => $row['list_id'],
                                        'cpn_type'     => $cpn_type,
                                        'cpn_id'       => $row['id'],
                                        'review_type'  => 3,
                                        'messages'     => '数据类型不正确',
                                        'row'          => $j,
                                        'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                                    ];
                                    continue;
                                } else
                                {
                                    if (strpos($numVal, '_') === false)
                                    {
                                        //不包含下划线
                                        //转换字符串
                                        if (strpos($numVal, '.') !== false)
                                        {
                                            $numVal = floatval($numVal);
                                        } else
                                        {
                                            $numVal = intval($numVal);
                                        }

                                        if ($numVal > 0)
                                        {
                                            if (($field !== 'cpn_ref_price') && !is_int($numVal))//供货周期
                                            {
                                                //问题行数递增
                                                if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                                {
                                                    $j++;
                                                }

                                                //更新该条元器件审查状态
                                                $batchData[$row['id']]['result_grc'] = 2;

                                                $errInfo[] = [
                                                    'main_task_id' => $main_task_id,
                                                    'list_id'      => $row['list_id'],
                                                    'cpn_type'     => $cpn_type,
                                                    'cpn_id'       => $row['id'],
                                                    'review_type'  => 3,
                                                    'messages'     => '数据类型不正确',
                                                    'row'          => $j,
                                                    'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                                                ];

                                                continue;
                                            }
                                        } else
                                        {
                                            //问题行数递增
                                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                            {
                                                $j++;
                                            }

                                            //更新该条元器件审查状态
                                            $batchData[$row['id']]['result_grc'] = 2;

                                            $errInfo[] = [
                                                'main_task_id' => $main_task_id,
                                                'list_id'      => $row['list_id'],
                                                'cpn_type'     => $cpn_type,
                                                'cpn_id'       => $row['id'],
                                                'review_type'  => 3,
                                                'messages'     => '数据类型不正确',
                                                'row'          => $j,
                                                'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                                            ];

                                            continue;
                                        }
                                    } else
                                    {
                                        //包含下划线
                                        $arr = explode('_', $numVal);
                                        if (count($arr) != 2)
                                        {
                                            //问题行数递增
                                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                            {
                                                $j++;
                                            }

                                            //更新该条元器件审查状态
                                            $batchData[$row['id']]['result_grc'] = 2;

                                            $errInfo[] = [
                                                'main_task_id' => $main_task_id,
                                                'list_id'      => $row['list_id'],
                                                'cpn_type'     => $cpn_type,
                                                'cpn_id'       => $row['id'],
                                                'review_type'  => 3,
                                                'messages'     => '数据类型不正确',
                                                'row'          => $j,
                                                'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                                            ];
                                            continue;
                                        }

                                        if ($arr[0] >= $arr[1])
                                        {
                                            //问题行数递增
                                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                            {
                                                $j++;
                                            }

                                            //更新该条元器件审查状态
                                            $batchData[$row['id']]['result_grc'] = 2;

                                            $errInfo[] = [
                                                'main_task_id' => $main_task_id,
                                                'list_id'      => $row['list_id'],
                                                'cpn_type'     => $cpn_type,
                                                'cpn_id'       => $row['id'],
                                                'review_type'  => 3,
                                                'messages'     => '数据类型不正确',
                                                'row'          => $j,
                                                'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                                            ];
                                            continue;
                                        }

                                        foreach ($arr as $kk => $vv)
                                        {
                                            //如果字符串有空值
                                            if (empty($vv))
                                            {
                                                //问题行数递增
                                                if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                                {
                                                    $j++;
                                                }

                                                //更新该条元器件审查状态
                                                $batchData[$row['id']]['result_grc'] = 2;

                                                $errInfo[] = [
                                                    'main_task_id' => $main_task_id,
                                                    'list_id'      => $row['list_id'],
                                                    'cpn_type'     => $cpn_type,
                                                    'cpn_id'       => $row['id'],
                                                    'review_type'  => 3,
                                                    'messages'     => '数据类型不正确',
                                                    'row'          => $j,
                                                    'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                                                ];
                                                continue 2;
                                            }

                                            //转换字符串
                                            if (strpos($vv, '.') !== false)
                                            {
                                                $vv = floatval($vv);
                                            } else
                                            {
                                                $vv = intval($vv);
                                            }

                                            if ($vv > 0)
                                            {
                                                if (($field !== 'cpn_ref_price') && !is_int($vv))//供货周期
                                                {
                                                    //问题行数递增
                                                    if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                                    {
                                                        $j++;
                                                    }

                                                    //更新该条元器件审查状态
                                                    $batchData[$row['id']]['result_grc'] = 2;

                                                    $errInfo[] = [
                                                        'main_task_id' => $main_task_id,
                                                        'list_id'      => $row['list_id'],
                                                        'cpn_type'     => $cpn_type,
                                                        'cpn_id'       => $row['id'],
                                                        'review_type'  => 3,
                                                        'messages'     => '数据类型不正确',
                                                        'row'          => $j,
                                                        'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                                                    ];

                                                    continue 2;
                                                }
                                            } else
                                            {
                                                //问题行数递增
                                                if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                                {
                                                    $j++;
                                                }

                                                //更新该条元器件审查状态
                                                $batchData[$row['id']]['result_grc'] = 2;

                                                $errInfo[] = [
                                                    'main_task_id' => $main_task_id,
                                                    'list_id'      => $row['list_id'],
                                                    'cpn_type'     => $cpn_type,
                                                    'cpn_id'       => $row['id'],
                                                    'review_type'  => 3,
                                                    'messages'     => '数据类型不正确',
                                                    'row'          => $j,
                                                    'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                                                ];

                                                continue 2;
                                            }
                                        }
                                    }
                                }
                            }

                        }
                        //                        //字典判断
                        if (in_array($field, $domesticRules['scope']['fields']) && (!empty($val) || ($val === "0")) && !in_array($val, $domesticRules['scope']['rules'][$field]))
                        {
                            //问题行数递增
                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                            {
                                $j++;
                            }

                            //更新该条元器件审查状态
                            if ($field == "cpn_is_core_important")
                            {
                                //更新该条元器件审查状态
                                $batchData[$row['id']]['yield_is_core_important'] = 1;
                                $batchData[$row['id']]['result_grc']              = 2;
                            }
                            if ($field == "cpn_control_level")
                            {
                                //更新该条元器件审查状态
                                $batchData[$row['id']]['yield_control_level'] = "E";
                                $batchData[$row['id']]['result_grc']          = 2;

                            }

                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $row['id'],
                                'review_type'  => 4,
                                'messages'     => '字典的字段超出范围',
                                'row'          => $j,
                                'msg_position' => $j . '行' . $domesticRules['header'][$field] . "列",
                            ];


                            $i++;
                            continue;

                        }


                    }
                }

                //必须先更新一次，后续才可正确判断一致性（如:自主可控等级出现F、G等超出范围值，则按最低等级的话，就出错了）

                //批量更新审查结果
                $domModel = new \APP\Model\CpnDomestic();

                $domModel->updateBatch($batchData);


                $listIds = array_unique(array_column($data, 'list_id'));

                $cpnData = CpnDomestic::selectRaw("id,concat(list_id , cpn_category_code , cpn_name , cpn_specification_model , equip_use_number , cpn_quality , cpn_manufacturer , cpn_package , cpn_control_level , cpn_is_core_important , cpn_ref_price , cpn_period, cpn_detect_apartment , history , remark) as cpn_info")
                    ->whereIn('list_id', $listIds)
                    ->get()->toArray();
                //数组分类
                $newData      = [];
                $batchNewData = [];
                foreach ($cpnData as $ks => $v)
                {
                    $newData[$v['cpn_info']][] = $v['id'];
                }
                //判断重复元器件
                foreach ($newData as $ks => $v)
                {
                    if (count($v) > 1)//证明有重复的数据
                    {
                        $msg_position = "";
                        for ($s = 0; $s < count($v); $s++)
                        {
                            $msg_position .= (++$j) . "行 ";
                        }
                        foreach ($v as $key => $val)
                        {
                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $val,
                                'review_type'  => 5,
                                'messages'     => '填报数据重复',
                                'row'          => $j,
                                'msg_position' => $msg_position,
                            ];
                            if (!empty($v[$key + 1]))
                            {
                                //更新该条元器件审查状态
                                $batchNewData[$v[$key + 1]]['id']         = $v[$key + 1];
                                $batchNewData[$v[$key + 1]]['is_repeat']  = 1;
                                $batchNewData[$v[$key + 1]]['result_grc'] = 2;
                            }
                        }
                        $i++;
                        continue;
                    }
                }


                //一致性校验
                $cpnData      = CpnDomestic::selectRaw("id , yield_control_level , yield_is_core_important , concat(list_id , cpn_specification_model , cpn_manufacturer) as cpn_info")
                    ->whereIn('list_id', $listIds)
                    ->get()->toArray();
                $cpnNewlDatas = [];
                foreach ($cpnData as $k => $v)
                {
                    $cpnNewlDatas[$v['cpn_info']][] = $v;
                }
                //                dd($cpnNewlDatas);
                foreach ($cpnNewlDatas as $ks => $v)
                {
                    //自主可控一致性判断
                    $controlNewData = $this->array_unique_fb($v, ['yield_control_level']);
                    if (count($controlNewData) > 1)//代表自主可控数据不一致
                    {
                        //获取最低自主可控等级
                        $controlDatas = array_column($v, 'yield_control_level');
                        arsort($controlDatas);
                        $smallControlLevel = reset($controlDatas);


                        $msg_position = "";
                        for ($s = 0; $s < count($v); $s++)
                        {
                            $msg_position .= (++$j) . "行 ";
                        }
                        foreach ($v as $key => $val)
                        {
                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $val['id'],
                                'review_type'  => 7,
                                'messages'     => '自主可控等级数据不一致',
                                'row'          => $j,
                                'msg_position' => $msg_position,
                            ];

                            //更新该条元器件审查状态
                            $batchNewData[$val['id']]['id']                  = $val['id'];
                            $batchNewData[$val['id']]['yield_control_level'] = $smallControlLevel;
                            $batchNewData[$val['id']]['result_grc']          = 2;
                        }

                        $i++;
                        continue;
                    }
                }


                //核心关键一致性判断
                foreach ($cpnNewlDatas as $ks => $v)
                {
                    $coreNewData = $this->array_unique_fb($v, ['yield_is_core_important']);
                    //                    dd($coreNewData);
                    if (count($coreNewData) > 1)//代表核心关键数据不一致
                    {
                        $msg_position = "";
                        for ($s = 0; $s < count($v); $s++)
                        {
                            $msg_position .= (++$j) . "行 ";
                        }
                        foreach ($v as $key => $val)
                        {
                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $val['id'],
                                'review_type'  => 6,
                                'messages'     => '核心关键数据不一致',
                                'row'          => $j,
                                'msg_position' => $msg_position,
                            ];

                            //更新该条元器件审查状态
                            $batchNewData[$val['id']]['id']                      = $val['id'];
                            $batchNewData[$val['id']]['yield_is_core_important'] = 1;
                            $batchNewData[$val['id']]['result_grc']              = 2;
                        }

                        $i++;
                        continue;
                    }
                }


                //循环构造相同字段数组值
                foreach ($batchNewData as $kl => $vl)
                {
                    if (!isset($vl['yield_control_level']))
                    {
                        $batchNewData[$kl]['yield_control_level'] = $batchData[$kl]['yield_control_level'];
                    }
                    if (!isset($vl['yield_is_core_important']))
                    {
                        $batchNewData[$kl]['yield_is_core_important'] = $batchData[$kl]['yield_is_core_important'];
                    }
                    if (!isset($vl['is_repeat']))
                    {
                        $batchNewData[$kl]['is_repeat'] = $batchData[$kl]['is_repeat'];
                    }
                    if (!isset($vl['result_grc']))
                    {
                        $batchNewData[$kl]['result_grc'] = $batchData[$kl]['result_grc'];
                    }
                }


                //批量更新审查结果
                $domModel = new \APP\Model\CpnDomestic();
                if (!empty($batchNewData))
                {
                    $domModel->updateBatch($batchNewData);
                }


                //更新主任务的已审查的元器件数量
                MainTask::where('id', $main_task_id)->increment('cpn_checked_num', count($rows));
                MainTask::where('id', $main_task_id)->increment('error_num', $i);
            }
        } elseif ($fileType == '进口')
        {
            $importRules  = $this->verifyImportRules;
            $importTitle3 = $this->importTitle3;

            $j = 1;//记录错误行数

            //分批审查
            $chunks = collect($data)->chunk($chunkSize);
            foreach ($chunks as $block)
            {
                $rows = $block->toArray();
                $i    = 0;
                foreach ($rows as $k => $row)
                {
                    //初始化数组
                    $batchImpData[$row['id']] = [
                        'id'                        => $row['id'],
                        "yield_is_core_important"   => @$row['cpn_is_core_important'],
                        "yield_safe_color"          => @$row['safe_color'],
                        "yield_proposed_safe_color" => @$row['proposed_safe_color'],
                        "yield_necessity"           => @$row['necessity'],
                        'is_repeat'                 => 0,
                        "result_grc"                => 3
                    ];
                    foreach ($row as $field => $val)
                    {

                        //判断关键信息必填项 是否为空
                        if (in_array($field, $importRules['nopass']) && (empty($val) && $val !== 0))
                        {
                            //问题行数递增
                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                            {
                                $j++;
                            }

                            $batchImpData[$row['id']]['result_grc'] = 1;

                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $row['id'],
                                'review_type'  => 1,
                                'messages'     => $importTitle3[$field].'信息缺失',
                                'row'          => $j,
                                'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                            ];

                            $i++;
                            break;

                        }



                        // 判断分类代码
                        if (in_array($field, $importRules['scope']['category']) && (!empty($val) || ($val === "0")) && !in_array($val, $category))
                        {
                            //问题行数递增
                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                            {
                                $j++;
                            }

                            //更新该条元器件审查状态
                            $batchImpData[$row['id']]['result_grc'] = 1;


                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $row['id'],
                                'review_type'  => 1,
                                'messages'     =>'分类代码不存在',
                                'row'          => $j,
                                'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                            ];

                            $i++;
                            break;
                        }

                        // 判断是否为空
                        if (in_array($field, $importRules['required']) && (empty($val) && $val !== 0))
                        {
                            //问题行数递增
                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                            {
                                $j++;
                            }

                            //更新该条元器件审查状态
                            if ($field == "cpn_is_core_important")
                            {
                                $batchImpData[$row['id']]['yield_is_core_important'] = 1;
                                $batchImpData[$row['id']]['result_grc']              = 2;

                            }

                            if ($field == "proposed_safe_color")
                            {
                                if (empty($row['safe_color']))
                                {
                                    $batchImpData[$row['id']]['yield_proposed_safe_color'] = "红色";
                                    $batchImpData[$row['id']]['result_grc']                = 2;

                                } else
                                {

                                    $batchImpData[$row['id']]['yield_proposed_safe_color'] = $row['safe_color'];
                                    $batchImpData[$row['id']]['result_grc']                = 2;
                                }
                            }
                            if ($field == "safe_color")
                            {
                                $batchImpData[$row['id']]['safe_color'] = "红色";
                                $batchImpData[$row['id']]['result_grc'] = 2;

                            }
                            if ($field == "necessity")
                            {
                                $batchImpData[$row['id']]['yield_necessity'] = 3;
                                $batchImpData[$row['id']]['result_grc']      = 2;
                            }

                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $row['id'],
                                'review_type'  => 2,
                                'messages'     => '内容不能为空',
                                'row'          => $j,
                                'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                            ];


                            $i++;
                            continue;
                        }


                        // 验证数据类型
                        if (in_array($field, $importRules['dataType']['fields']) && ($val != ''))
                        {
                            //转化成数值型
                            $numVal = $val;
                            if (is_string($numVal))
                            {
                                //字符串必须由数字和.组成
                                if (preg_match('/[^0-9\.]/', $numVal))
                                {
                                    //问题行数递增
                                    if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                    {
                                        $j++;
                                    }

                                    //更新该条元器件审查状态
                                    $batchImpData[$row['id']]['result_grc'] = 2;

                                    $errInfo[] = [
                                        'main_task_id' => $main_task_id,
                                        'list_id'      => $row['list_id'],
                                        'cpn_type'     => $cpn_type,
                                        'cpn_id'       => $row['id'],
                                        'review_type'  => 3,
                                        'messages'     => '数据类型不正确',
                                        'row'          => $j,
                                        'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                                    ];

                                    $i++;
                                    continue;
                                }
                                if (strpos($numVal, '.') !== false)
                                {
                                    $numVal = floatval($numVal);
                                } else
                                {
                                    $numVal = intval($numVal);
                                }
                            }

                            if ($numVal > 0)
                            {
                                if (($field !== 'cpn_ref_price') && !($field == 'equip_use_number' && in_array(substr($row['cpn_category_code'], 0, 4), $this->LineCategories)) && ($field !== 'cpn_period') && !is_int($numVal))//参考价格
                                {
                                    //问题行数递增
                                    if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                    {
                                        $j++;
                                    }

                                    //更新该条元器件审查状态
                                    $batchImpData[$row['id']]['result_grc'] = 2;

                                    $errInfo[] = [
                                        'main_task_id' => $main_task_id,
                                        'list_id'      => $row['list_id'],
                                        'cpn_type'     => $cpn_type,
                                        'cpn_id'       => $row['id'],
                                        'review_type'  => 3,
                                        'messages'     => '数据类型不正确',
                                        'row'          => $j,
                                        'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                                    ];

                                    $i++;
                                    continue;
                                }
                            } else
                            {
                                //问题行数递增
                                if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                {
                                    $j++;
                                }

                                //更新该条元器件审查状态
                                $batchImpData[$row['id']]['result_grc'] = 2;

                                $errInfo[] = [
                                    'main_task_id' => $main_task_id,
                                    'list_id'      => $row['list_id'],
                                    'cpn_type'     => $cpn_type,
                                    'cpn_id'       => $row['id'],
                                    'review_type'  => 3,
                                    'messages'     => '数据类型不正确',
                                    'row'          => $j,
                                    'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                                ];

                                $i++;
                                continue;
                            }

                        }

                        // 验证数据格式
                        if (in_array($field, $importRules['dataFormat']) && ($val != ''))
                        {
                            //转化成数值型
                            $numVal = $val;
                            if (is_string($numVal))
                            {
                                //字符串必须由数字和. _组成
                                if (preg_match('/[^0-9\._]/', $numVal))
                                {
                                    //问题行数递增
                                    if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                    {
                                        $j++;
                                    }

                                    //更新该条元器件审查状态
                                    $batchImpData[$row['id']]['result_grc'] = 2;

                                    $errInfo[] = [
                                        'main_task_id' => $main_task_id,
                                        'list_id'      => $row['list_id'],
                                        'cpn_type'     => $cpn_type,
                                        'cpn_id'       => $row['id'],
                                        'review_type'  => 3,
                                        'messages'     => '数据类型不正确',
                                        'row'          => $j,
                                        'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                                    ];

                                    continue;
                                } else
                                {
                                    if (strpos($numVal, '_') === false)
                                    {
                                        //不包含下划线
                                        //转换字符串
                                        if (strpos($numVal, '.') !== false)
                                        {
                                            $numVal = floatval($numVal);
                                        } else
                                        {
                                            $numVal = intval($numVal);
                                        }

                                        if ($numVal > 0)
                                        {
                                            if (($field !== 'cpn_ref_price') && !is_int($numVal))//供货周期
                                            {
                                                //问题行数递增
                                                if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                                {
                                                    $j++;
                                                }

                                                //更新该条元器件审查状态
                                                $batchImpData[$row['id']]['result_grc'] = 2;

                                                $errInfo[] = [
                                                    'main_task_id' => $main_task_id,
                                                    'list_id'      => $row['list_id'],
                                                    'cpn_type'     => $cpn_type,
                                                    'cpn_id'       => $row['id'],
                                                    'review_type'  => 3,
                                                    'messages'     => '数据类型不正确',
                                                    'row'          => $j,
                                                    'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                                                ];

                                                continue;
                                            }
                                        } else
                                        {
                                            //问题行数递增
                                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                            {
                                                $j++;
                                            }

                                            //更新该条元器件审查状态
                                            $batchImpData[$row['id']]['result_grc'] = 2;

                                            $errInfo[] = [
                                                'main_task_id' => $main_task_id,
                                                'list_id'      => $row['list_id'],
                                                'cpn_type'     => $cpn_type,
                                                'cpn_id'       => $row['id'],
                                                'review_type'  => 3,
                                                'messages'     => '数据类型不正确',
                                                'row'          => $j,
                                                'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                                            ];


                                            continue;
                                        }
                                    } else
                                    {
                                        //包含下划线
                                        $arr = explode('_', $numVal);
                                        if (count($arr) != 2)
                                        {
                                            //问题行数递增
                                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                            {
                                                $j++;
                                            }

                                            //更新该条元器件审查状态
                                            $batchImpData[$row['id']]['result_grc'] = 2;

                                            $errInfo[] = [
                                                'main_task_id' => $main_task_id,
                                                'list_id'      => $row['list_id'],
                                                'cpn_type'     => $cpn_type,
                                                'cpn_id'       => $row['id'],
                                                'review_type'  => 3,
                                                'messages'     => '数据类型不正确',
                                                'row'          => $j,
                                                'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                                            ];

                                            continue;
                                        }

                                        if ($arr[0] >= $arr[1])
                                        {
                                            //问题行数递增
                                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                            {
                                                $j++;
                                            }

                                            //更新该条元器件审查状态
                                            $batchImpData[$row['id']]['result_grc'] = 2;

                                            $errInfo[] = [
                                                'main_task_id' => $main_task_id,
                                                'list_id'      => $row['list_id'],
                                                'cpn_type'     => $cpn_type,
                                                'cpn_id'       => $row['id'],
                                                'review_type'  => 3,
                                                'messages'     => '数据类型不正确',
                                                'row'          => $j,
                                                'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                                            ];

                                            continue;
                                        }

                                        foreach ($arr as $kk => $vv)
                                        {
                                            //如果字符串有空值
                                            if (empty($vv))
                                            {
                                                //问题行数递增
                                                if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                                {
                                                    $j++;
                                                }

                                                //更新该条元器件审查状态
                                                $batchImpData[$row['id']]['result_grc'] = 2;

                                                $errInfo[] = [
                                                    'main_task_id' => $main_task_id,
                                                    'list_id'      => $row['list_id'],
                                                    'cpn_type'     => $cpn_type,
                                                    'cpn_id'       => $row['id'],
                                                    'review_type'  => 3,
                                                    'messages'     => '数据类型不正确',
                                                    'row'          => $j,
                                                    'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                                                ];

                                                continue 2;
                                            }

                                            //转换字符串
                                            if (strpos($vv, '.') !== false)
                                            {
                                                $vv = floatval($vv);
                                            } else
                                            {
                                                $vv = intval($vv);
                                            }

                                            if ($vv > 0)
                                            {
                                                if (($field !== 'cpn_ref_price') && !is_int($vv))//供货周期
                                                {
                                                    //问题行数递增
                                                    if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                                    {
                                                        $j++;
                                                    }

                                                    //更新该条元器件审查状态
                                                    $batchImpData[$row['id']]['result_grc'] = 2;

                                                    $errInfo[] = [
                                                        'main_task_id' => $main_task_id,
                                                        'list_id'      => $row['list_id'],
                                                        'cpn_type'     => $cpn_type,
                                                        'cpn_id'       => $row['id'],
                                                        'review_type'  => 3,
                                                        'messages'     => '数据类型不正确',
                                                        'row'          => $j,
                                                        'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                                                    ];


                                                    continue 2;
                                                }
                                            } else
                                            {
                                                //问题行数递增
                                                if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                                                {
                                                    $j++;
                                                }

                                                //更新该条元器件审查状态
                                                $batchImpData[$row['id']]['result_grc'] = 2;

                                                $errInfo[] = [
                                                    'main_task_id' => $main_task_id,
                                                    'list_id'      => $row['list_id'],
                                                    'cpn_type'     => $cpn_type,
                                                    'cpn_id'       => $row['id'],
                                                    'review_type'  => 3,
                                                    'messages'     => '数据类型不正确',
                                                    'row'          => $j,
                                                    'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                                                ];

                                                continue 2;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        //字典判断
                        if (in_array($field, $importRules['scope']['fields']) && (!empty($val) || ($val === "0")) && !in_array($val, $importRules['scope']['rules'][$field]))
                        {
                            //问题行数递增
                            if (!in_array($row['id'], array_column($errInfo, 'cpn_id')))
                            {
                                $j++;
                            }

                            //更新该条元器件审查状态
                            if ($field == "cpn_is_core_important")
                            {
                                //更新该条元器件审查状态
                                $batchImpData[$row['id']]['cpn_is_core_important'] = 1;
                                $batchImpData[$row['id']]['result_grc']            = 2;

                            }
                            if ($field == "safe_color")
                            {
                                //更新该条元器件审查状态
                                $batchImpData[$row['id']]['yield_safe_color'] = "红色";
                                $batchImpData[$row['id']]['result_grc']       = 2;

                            }
                            if ($field == "proposed_safe_color")
                            {
                                if (empty($row['safe_color']))
                                {
                                    //更新该条元器件审查状态
                                    $batchImpData[$row['id']]['yield_proposed_safe_color'] = "红色";
                                    $batchImpData[$row['id']]['result_grc']                = 2;

                                } else
                                {
                                    //更新该条元器件审查状态
                                    $batchImpData[$row['id']]['yield_proposed_safe_color'] = $row['safe_color'];
                                    $batchImpData[$row['id']]['result_grc']                = 2;
                                }
                            }
                            if ($field == "necessity")
                            {
                                //更新该条元器件审查状态
                                $batchImpData[$row['id']]['yield_necessity'] = 3;
                                $batchImpData[$row['id']]['result_grc']      = 2;

                            }

                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $row['id'],
                                'review_type'  => 4,
                                'messages'     => '字典的字段超出范围',
                                'row'          => $j,
                                'msg_position' => $j . '行' . $importRules['header'][$field] . "列",
                            ];

                            $i++;
                            continue;
                        }

                    }
                }



                //必须先更新一次，后续才可正确判断一致性（如:自主可控等级出现F、G等超出范围值，则按最低等级的话，就出错了）

                //批量更新审查结果
                $impModel = new \APP\Model\CpnImport();

                $impModel->updateBatch($batchImpData);

                //数据重复校验
                $listIds = array_unique(array_column($data, 'list_id'));

                $cpnData = CpnImport::selectRaw("id,concat(list_id , cpn_category_code , cpn_name , cpn_specification_model , equip_use_number , cpn_quality , cpn_country , cpn_manufacturer , cpn_package , safe_color , proposed_safe_color , cpn_is_core_important , cpn_ref_price , necessity , cpn_period, access_channel , remark) as cpn_info")
                    ->whereIn('list_id', $listIds)
                    ->get()->toArray();

                //数组分类
                $newData         = [];
                $batchNewImpData = [];
                foreach ($cpnData as $ks => $v)
                {
                    $newData[$v['cpn_info']][] = $v['id'];
                }
                //判断重复元器件
                foreach ($newData as $ks => $v)
                {
                    if (count($v) > 1)//证明有重复的数据
                    {
                        $msg_position = "";
                        for ($s = 0; $s < count($v); $s++)
                        {
                            $msg_position .= (++$j) . "行 ";
                        }
                        foreach ($v as $key => $val)
                        {
                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $val,
                                'review_type'  => 5,
                                'messages'     => '填报数据重复',
                                'row'          => $j,
                                'msg_position' => $msg_position,
                            ];
                            if (!empty($v[$key + 1]))
                            {
                                //更新该条元器件审查状态
                                $batchNewImpData[$v[$key + 1]]['id']         = $v[$key + 1];
                                $batchNewImpData[$v[$key + 1]]['is_repeat']  = 1;
                                $batchNewImpData[$v[$key + 1]]['result_grc'] = 2;
                            }
                        }
                        $i++;
                        continue;
                    }
                }


                //一致性校验
                $cpnData      = CpnImport::selectRaw("id , yield_is_core_important , yield_safe_color, yield_proposed_safe_color , concat(list_id , cpn_specification_model) as cpn_info")
                    ->whereIn('list_id', $listIds)
                    ->get()->toArray();
                $cpnNewlDatas = [];
                foreach ($cpnData as $k => $v)
                {
                    $cpnNewlDatas[$v['cpn_info']][] = $v;
                }


                //核心关键一致性判断
                foreach ($cpnNewlDatas as $ks => $v)
                {
                    $coreNewData = $this->array_unique_fb($v, ['yield_is_core_important']);
                    //                                        dd($coreNewData);
                    if (count($coreNewData) > 1)//代表核心关键数据不一致
                    {
                        $msg_position = "";
                        for ($s = 0; $s < count($v); $s++)
                        {
                            $msg_position .= (++$j) . "行 ";
                        }
                        foreach ($v as $key => $val)
                        {
                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $val['id'],
                                'review_type'  => 6,
                                'messages'     => '核心关键数据不一致',
                                'row'          => $j,
                                'msg_position' => $msg_position,
                            ];

                            //更新该条元器件审查状态
                            $batchNewImpData[$val['id']]['id']                      = $val['id'];
                            $batchNewImpData[$val['id']]['yield_is_core_important'] = 1;
                            $batchNewImpData[$val['id']]['result_grc']              = 2;

                        }

                        $i++;
                        continue;
                    }
                }


                $safe_color_level = ['红色' => 5, "紫色" => 4, "橙色" => 3, "黄色" => 2, "绿色" => 1];//加严等级由严到松

                //安全颜色等级一致性判断
                foreach ($cpnNewlDatas as $k => $v)
                {
                    $colorNewData = $this->array_unique_fb($v, ['yield_safe_color']);

                    if (count($colorNewData) > 1)//代表安全颜色等级数据不一致
                    {
                        //找出最严的安全颜色等级
                        $colors        = array_column($v, 'yield_safe_color');
                        $newColorLevel = [];
                        foreach ($colors as $ki => $vs)
                        {
                            if(in_array($vs , $safe_color_level))
                            {
                                $newColorLevel[$safe_color_level[$vs]] = $vs;
                            }
                        }
                        krsort($newColorLevel);
                        $yanColor = reset($newColorLevel);


                        $msg_position = "";
                        for ($s = 0; $s < count($v); $s++)
                        {
                            $msg_position .= (++$j) . "行 ";
                        }
                        foreach ($v as $key => $val)
                        {
                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $val['id'],
                                'review_type'  => 8,
                                'messages'     => '安全等级颜色数据不一致',
                                'row'          => $j,
                                'msg_position' => $msg_position,
                            ];

                            //更新该条元器件审查状态
                            $batchNewImpData[$val['id']]['id']               = $val['id'];
                            $batchNewImpData[$val['id']]['yield_safe_color'] = $yanColor;
                            $batchNewImpData[$val['id']]['result_grc']       = 2;
                        }

                        $i++;
                        continue;
                    }
                }

                //建议安全颜色等级一致性判断
                foreach ($cpnNewlDatas as $k => $v)
                {
                    $pcolorNewData = $this->array_unique_fb($v, ['yield_proposed_safe_color']);

                    if (count($pcolorNewData) > 1)//代表安全颜色等级数据不一致
                    {
                        //找出最严的安全颜色等级
                        $colors        = array_column($v, 'yield_proposed_safe_color');
                        $newColorLevel = [];
                        foreach ($colors as $ki => $vs)
                        {
                            if(in_array($vs , $safe_color_level))
                            {
                                $newColorLevel[$safe_color_level[$vs]] = $vs;
                            }
                        }
                        krsort($newColorLevel);
                        $yanColor = reset($newColorLevel);


                        $msg_position = "";
                        for ($s = 0; $s < count($v); $s++)
                        {
                            $msg_position .= (++$j) . "行 ";
                        }
                        foreach ($v as $key => $val)
                        {
                            $errInfo[] = [
                                'main_task_id' => $main_task_id,
                                'list_id'      => $row['list_id'],
                                'cpn_type'     => $cpn_type,
                                'cpn_id'       => $val['id'],
                                'review_type'  => 9,
                                'messages'     => '建议安全等级颜色数据不一致',
                                'row'          => $j,
                                'msg_position' => $msg_position,
                            ];

                            //更新该条元器件审查状态
                            $batchNewImpData[$val['id']]['id']               = $val['id'];
                            $batchNewImpData[$val['id']]['yield_safe_color'] = $yanColor;
                            $batchNewImpData[$val['id']]['result_grc']       = 2;
                        }

                        $i++;
                        continue;
                    }
                }


                //循环构造相同字段数组值
                foreach ($batchNewImpData as $kl => $vl)
                {
                    if (!isset($vl['yield_is_core_important']))
                    {
                        $batchNewImpData[$kl]['yield_is_core_important'] = $batchImpData[$kl]['yield_is_core_important'];
                    }
                    if (!isset($vl['yield_safe_color']))
                    {
                        $batchNewImpData[$kl]['yield_safe_color'] = $batchImpData[$kl]['yield_safe_color'];
                    }
                    if (!isset($vl['yield_proposed_safe_color']))
                    {
                        $batchNewImpData[$kl]['yield_proposed_safe_color'] = $batchImpData[$kl]['yield_proposed_safe_color'];
                    }
                    if (!isset($vl['yield_necessity']))
                    {
                        $batchNewImpData[$kl]['yield_necessity'] = $batchImpData[$kl]['yield_necessity'];
                    }
                    if (!isset($vl['is_repeat']))
                    {
                        $batchNewImpData[$kl]['is_repeat'] = $batchImpData[$kl]['is_repeat'];
                    }
                    if (!isset($vl['result_grc']))
                    {
                        $batchNewImpData[$kl]['result_grc'] = $batchImpData[$kl]['result_grc'];
                    }
                }


                //批量更新审查结果
                $impModel = new \APP\Model\CpnImport();

                if (!empty($batchNewImpData))
                {
                    $impModel->updateBatch($batchNewImpData);
                }


                //                dd($i , $diffCpnIds , $errInfo, $batchImpCoreData , $batchImpStatusData , $batchImpRepeatData , $batchImpNecessityData , $batchImpPSafeColorData , $batchImpSafeColorData);


                //更新主任务的已审查的元器件数量
                MainTask::where('id', $main_task_id)->increment('cpn_checked_num', count($rows));
                MainTask::where('id', $main_task_id)->increment('error_num', $i);
            }
        }

        return $errInfo;
    }

    /**
     * 辅助审查 - 校验数据 - 批量
     * @param $main_task_id
     * @param $reviewList1
     * @param $reviewList2
     * @param $ruleData
     * @return array
     * @throws \Exception
     */
    public function verifyDataByAuxiliaryBulk($main_task_id, $reviewList1, $reviewList2, $ruleData)
    {
        if (empty($main_task_id) || empty($ruleData))
            return [];

        if (!$mainTask = MainTask::find($main_task_id))
            return [];

        $resultData = [];
        $ruleIds    = array_column($ruleData, 'rule_id');
        foreach ($ruleData as $key => $rule)
        {
            if (in_array($rule['rule_id'], [9, 10, 11, 12]))
                continue;

            $errInfo = [];
            switch ($rule['rule_id'])
            {
                case 1:
                    $errInfo = $this->crossVersion($main_task_id, $mainTask['model_unique_code'], $reviewList1, $reviewList2, $rule['contrast_main_task_id']);//跨版本比对审查
                    break;
                case 2:
                    $errInfo = $this->adoptDomesticReplace($main_task_id, $mainTask['model_unique_code'], $reviewList1, $reviewList2);//研制单位反馈采纳国产化替代数据
                    break;
                case 3:
                    $errInfo = $this->implementFake($main_task_id, $mainTask['model_unique_code'], $reviewList1);//研制单位反馈存在伪空包现象的数据
                    break;
                case 4:
                    $errInfo = $this->fakeEmpty($main_task_id, $reviewList1);//伪国产化、空心国产化、包装国产化知识库
                    break;
                case 5:
                    $errInfo = $this->historyFakeEmpty($main_task_id, $reviewList1);//历史伪空包现象专家审查数据
                    break;
                case 6:
                    $errInfo = $this->historySatisfy($main_task_id, $reviewList1);//历史满足度专家审查数据
                    break;
                case 7:
                    $this->ycd($main_task_id, $reviewList2);//对进口产品（技术）YCD数据（电子元器件类）
                    break;
                case 8:
                    $this->cots($main_task_id, $reviewList2);//COTS类器件质量等级字典
                    break;
                case 13:
                    $errInfo = $this->historySecurity($main_task_id, $reviewList2);//历史安全性审查专家审查数据
                    break;
                case 14:
                    $errInfo = $this->historyEnsure($main_task_id, $reviewList2);//历史可保障性审查专家审查数据
                    break;
                case 15:
                    $errInfo = $this->crossStage($main_task_id, $reviewList2, $rule['contrast_main_task_id']);//进口清单跨阶段比对审查
                    break;
                default:

                    break;
            }

            if (!empty($errInfo))
                $resultData = array_merge($resultData, $errInfo);
        }

        //进口清单国产化替代
        if ($isReplace = array_intersect($ruleIds, [9, 10, 11]))
        {
            $replaceData = $this->domesticReplace($main_task_id, $reviewList2, $ruleIds);
        } else
        {
            $replaceData = [];
        }

        //安全性、可保障性审查
        if ($isColor = array_intersect($ruleIds, [13, 14]))
        {
            $listIds     = CpnFiles::where('main_task_id', $main_task_id)->pluck('id')->toArray();
            $reviewList2 = CpnImport::whereIn('list_id', $listIds)
                ->selectRaw('id, list_id, cpn_specification_model, cpn_manufacturer, group_concat(safe_color separator "、") as safe_color')
                ->groupBy('cpn_specification_model', 'cpn_manufacturer')
                ->orderBy('id', 'asc')
                ->get()
                ->toArray(); //进口数据
            $colorData   = $this->doColor($main_task_id, $reviewList2);
        } else
        {
            $colorData = [];
        }

        return [
            'data'         => $resultData,//审查结果
            'replace_data' => $replaceData,//国产替代结果
            'color_data'   => $colorData//安全性、可保障性审查结果
        ];
    }

    /**
     * 辅助审查 - 批量
     * @param $main_task_id
     * @return array
     * @throws \Exception
     */
    public function reviewData($main_task_id)
    {
        if (empty($main_task_id))
            return [];

        if (!$mainTask = MainTask::find($main_task_id))
            return [];

        //1.YCD审查
        $listIds = CpnFiles::where('main_task_id', $main_task_id)->pluck('id')->toArray();
        $resData = $this->domesticReplace($main_task_id, $listIds);
        MainTask::where('id', $main_task_id)->increment('auxiliary_checked_num'); //更新完成任务数

        //2.意见判定
        $opinionData = !empty($resData) ? $this->getOpinionData($resData['newData']) : [];
        MainTask::where('id', $main_task_id)->increment('auxiliary_checked_num'); //更新完成任务数

        //3.伪空包推荐
        $recommendData = $this->getFakeData($main_task_id, $listIds);
        MainTask::where('id', $main_task_id)->increment('auxiliary_checked_num'); //更新完成任务数

        return [
            'opinionData'   => $opinionData,//审查结果
            'replaceData'   => $resData['replaceData'],//国产替代结果
            'recommendData' => $recommendData
        ];
    }

    /**
     * 国产军用电子元器件手册
     * @param $rows
     * @param $main_task_id
     * @return array
     */
    public function getResultByBulk3($rows, $main_task_id)
    {
        return true;
    }

    /**
     * 赛思库®电子元器件国产化替代数据
     * @param $rows
     * @param $main_task_id
     * @param $is_history
     * @return array
     */
    public function getResultByBulk4($rows, $main_task_id)
    {
        return true;
    }

    /**
     * 赛思库®电子元器件性能指标参数数据
     * @param $rows
     * @param $main_task_id
     * @param $model
     * @param $is_history
     * @return array
     */
    public function getResultByBulk5($rows, $main_task_id, $model)
    {
        if (empty($rows) || empty($main_task_id) || empty($model))
            return [];

        $errors = [];
        foreach ($rows as $key => $val)
        {
            $result = $model->getSimilarComponents($val['cpn_specification_model']);
            if (!empty($result['replace_cpns']))
            {
                foreach ($result['replace_cpns'] as $k => $v)
                {
                    //可替代产品应用经历
                    $experience = $this->getExperience($v['part_number']);
                    //历史可替代产品审查情况
                    $history  = !empty($is_history) ? $this->getHistory($v['part_number']) : '';
                    $errors[] = [
                        'main_task_id'                    => $main_task_id,
                        'list_id'                         => $val['list_id'],
                        'cpn_type'                        => 1,
                        'cpn_id'                          => $val['id'],
                        'rule_id'                         => 11,
                        'repalce_type'                    => 2,
                        'replace_cpn_specification_model' => $v['part_number'],
                        'replace_cpn_manufacturer'        => $v['part_manufacturer'],
                        'replace_cpn_quality'             => $v['quality_level'],
                        'experience'                      => $experience,
                        'history'                         => $history,
                    ];
                }

                unset($rows[$key]);
            }
        }

        $data = [
            'errors'      => $errors, //错误信息
            'uncompleted' => $rows, //未查出的数据
        ];

        return $data;
    }

    /**
     * 进口元器件 - 无匹配结果
     * @param $rows
     * @param $main_task_id
     * @return array
     */
    public function getResultByBulk6($rows, $main_task_id)
    {
        $data = [];
        foreach ($rows as $key => $val)
        {
            $data[] = [
                'main_task_id'                    => $main_task_id,
                'list_id'                         => $val['list_id'],
                'cpn_type'                        => 1,
                'cpn_id'                          => $val['id'],
                'rule_id'                         => '',
                'repalce_type'                    => 0,
                'replace_cpn_specification_model' => '',
                'replace_cpn_manufacturer'        => '',
                'replace_cpn_quality'             => '',
                'experience'                      => '',
                'history'                         => '',
            ];
        }

        return $data;
    }

    /**
     * 获取可替代产品应用经历
     * @param null $cpn_specification_model
     * @return string
     */
    public function getExperience($cpn_specification_model = null)
    {
        if (empty($cpn_specification_model))
            return '';

        if ($cpnNum = CpnDomestic::where('cpn_specification_model', $cpn_specification_model)->sum('equip_use_number'))
        {
            //型号数
            $listIds      = CpnDomestic::where('cpn_specification_model', $cpn_specification_model)->pluck('list_id')->toArray();
            $modelIds     = CpnFiles::whereIn('id', $listIds)->pluck('model_id')->toArray();
            $equipmentIds = ModelStructure::selectRaw('distinct equipment_id')->whereIn('id', $modelIds)->pluck('equipment_id')->toArray();
            $equipmentNum = count($equipmentIds);
            //装备数
            $equipmentTypeNum = Equipment::selectRaw('distinct type_id')->whereIn('id', $equipmentIds)->pluck('type_id')->count();
            $str              = "{$cpn_specification_model}规格型号在{$equipmentTypeNum}类个装备{$equipmentNum}个型号应用{$cpnNum}只";
            return $str;
        }

        return '';
    }

    /**
     * 获取可替代产品审查情况
     * @param null $cpn_specification_model
     * @return string
     */
    public function getHistory($cpn_specification_model = null)
    {
        if (empty($cpn_specification_model))
            return '';

        $str = '';
        //认可
        $acceptData = ProReviewOpinionImport::from('pro_review_opinion_import as i')
            ->leftJoin('equipment as e', 'e.model_unique_code', 'i.model_unique_code')
            ->leftJoin('equipment_type as et', 'et.id', 'e.type_id')
            ->select('i.expert_name', 'i.model_unique_code', 'et.name')
            ->where('i.kb_substitution_plan', $cpn_specification_model)
            ->groupBy('et.name', 'i.expert_name')
            ->get()
            ->toArray();
        if (!empty($acceptData))
        {
            $modelTypes = array_unique(array_column($acceptData, 'name'));
            foreach ($modelTypes as $key => $val)
            {
                $expertArr = [];
                foreach ($acceptData as $k => $v)
                {
                    if ($val == $v['name'])
                    {
                        $expertArr[] = '专家' . $v['expert_name'];
                    }
                }
                $str .= implode('、', $expertArr) . "在{$val}类型装备上认可了该替代方案；";
            }
        }

        return $str;
    }

    /**
     * 跨版本比对审查
     * @param $main_task_id
     * @param $model_unique_code
     * @param $reviewList1
     * @param $reviewList2
     * @param $contrast_main_task_id
     * @return array
     */
    public function crossVersion($main_task_id = null, $model_unique_code = null, $reviewList1 = [], $reviewList2 = [], $contrast_main_task_id = [])
    {
        if (empty($main_task_id) || empty($model_unique_code) || empty($contrast_main_task_id))
            return [];


        $data = [];//审查结果数据
        //1.国产清单比对
        if (!empty($reviewList1))
        {
            //对比版本数据
            $cmpFilesIds     = CpnFiles::where('main_task_id', $contrast_main_task_id)
                ->where('type', '国产')
                ->pluck('id')
                ->toArray();
            $cmpDomestic     = CpnDomestic::whereIn('list_id', $cmpFilesIds)
                ->selectRaw('id, list_id, cpn_specification_model, cpn_manufacturer, concat(cpn_specification_model,cpn_manufacturer) as spe_manufacturer')
                ->get()
                ->toArray();
            $cmpDomesticData = array_column($cmpDomestic, 'spe_manufacturer');

            //当前版本数据
            $domesticData = array_column($reviewList1, 'spe_manufacturer');

            //数据源
            $source     = ProReviewOpinionCloseDom::where('model_unique_code', $model_unique_code)
                ->where('is_agree_empty_package', '是')
                ->selectRaw('concat(cpn_specification_model,cpn_manufacturer) as spe_manufacturer')
                ->get()
                ->toArray();
            $sourceData = array_column($source, 'spe_manufacturer');

            $noPass = [];//审查不通过数据
            foreach ($reviewList1 as $key => $val)
            {
                if (in_array($val['spe_manufacturer'], $sourceData))
                {
                    $data[] = [
                        'main_task_id' => $main_task_id,
                        'list_id'      => $val['list_id'],
                        'cpn_type'     => 0,
                        'cpn_id'       => $val['id'],
                        'rule_id'      => 1,
                        'messages'     => '审查不通过',
                    ];

                    $noPass[] = $val['spe_manufacturer'];//审查不通过数据
                    unset($reviewList1[$key]);
                }
            }

            //处理剩余元器件
            foreach ($reviewList1 as $key => $val)
            {
                //无需专家审查
                if (in_array($val['spe_manufacturer'], $cmpDomesticData))
                {
                    $data[] = [
                        'main_task_id' => $main_task_id,
                        'list_id'      => $val['list_id'],
                        'cpn_type'     => 0,
                        'cpn_id'       => $val['id'],
                        'rule_id'      => 1,
                        'messages'     => '无需专家审查',
                    ];
                } else
                {//需进一步审查
                    $data[] = [
                        'main_task_id' => $main_task_id,
                        'list_id'      => $val['list_id'],
                        'cpn_type'     => 0,
                        'cpn_id'       => $val['id'],
                        'rule_id'      => 1,
                        'messages'     => '需进一步审查',
                    ];
                }
            }

            //统计减少的数据
            foreach ($cmpDomestic as $key => $val)
            {
                if (!in_array($val['spe_manufacturer'], $domesticData) && !in_array($val['spe_manufacturer'], $noPass))
                {
                    $data[] = [
                        'main_task_id' => $main_task_id,
                        'list_id'      => $val['list_id'],
                        'cpn_type'     => 0,
                        'cpn_id'       => $val['id'],
                        'rule_id'      => 1,
                        'messages'     => '比对减少数据',
                    ];
                }
            }
        }

        //2.进口清单比对
        if (!empty($reviewList2))
        {
            //获取对比版本清单信息
            $cmpFilesIds   = CpnFiles::where('main_task_id', $contrast_main_task_id)
                ->where('type', '进口')
                ->pluck('id')
                ->toArray();
            $cmpImport     = CpnImport::whereIn('list_id', $cmpFilesIds)
                ->selectRaw('id, list_id, cpn_specification_model, cpn_manufacturer, concat(cpn_specification_model,cpn_manufacturer) as spe_manufacturer')
                ->get()
                ->toArray();
            $cmpImportData = array_column($cmpImport, 'spe_manufacturer');

            //当前版本数据
            $importData = array_column($reviewList2, 'spe_manufacturer');

            //数据源
            $source     = ProReviewOpinionCloseImp::where('model_unique_code', $model_unique_code)
                ->where('is_agree_substitution', '是')
                ->selectRaw('concat(cpn_specification_model,cpn_manufacturer) as spe_manufacturer')
                ->get()
                ->toArray();
            $sourceData = array_column($source, 'spe_manufacturer');

            $noPass = [];//审查不通过数据
            foreach ($reviewList2 as $key => $val)
            {
                if (in_array($val['spe_manufacturer'], $sourceData))
                {
                    $data[] = [
                        'main_task_id' => $main_task_id,
                        'list_id'      => $val['list_id'],
                        'cpn_type'     => 1,
                        'cpn_id'       => $val['id'],
                        'rule_id'      => 1,
                        'messages'     => '审查不通过',
                    ];

                    $noPass[] = $val['spe_manufacturer'];
                    unset($reviewList2[$key]);
                }
            }

            //处理剩余元器件
            foreach ($reviewList2 as $key => $val)
            {
                //无需专家审查
                if (in_array($val['spe_manufacturer'], $cmpImportData))
                {
                    $data[] = [
                        'main_task_id' => $main_task_id,
                        'list_id'      => $val['list_id'],
                        'cpn_type'     => 1,
                        'cpn_id'       => $val['id'],
                        'rule_id'      => 1,
                        'messages'     => '无需专家审查',
                    ];
                } else
                {//需进一步审查
                    $data[] = [
                        'main_task_id' => $main_task_id,
                        'list_id'      => $val['list_id'],
                        'cpn_type'     => 1,
                        'cpn_id'       => $val['id'],
                        'rule_id'      => 1,
                        'messages'     => '需进一步审查',
                    ];
                }
            }

            //统计减少的数据
            foreach ($cmpImport as $key => $val)
            {
                if (!in_array($val['spe_manufacturer'], $importData) && !in_array($val['spe_manufacturer'], $noPass))
                {
                    $data[] = [
                        'main_task_id' => $main_task_id,
                        'list_id'      => $val['list_id'],
                        'cpn_type'     => 1,
                        'cpn_id'       => $val['id'],
                        'rule_id'      => 1,
                        'messages'     => '比对减少数据',
                    ];
                }
            }
        }

        //更新完成任务数
        MainTask::where('id', $main_task_id)->increment('auxiliary_checked_num');

        return $data;
    }

    /**
     * 研制单位反馈采纳国产化替代数据
     * @param $main_task_id
     * @param $model_unique_code
     * @param $reviewList1
     * @param $reviewList2
     * @return array
     */
    public function adoptDomesticReplace($main_task_id = null, $model_unique_code = null, $reviewList1 = [], $reviewList2 = [])
    {
        if (empty($main_task_id) || empty($model_unique_code))
            return [];

        $data         = [];//审查结果数据
        $domesticData = array_column($reviewList1, 'spe_manufacturer');//国产清单数据
        $importData   = array_column($reviewList2, 'spe_manufacturer');//进口清单数据
        $source       = ProReviewOpinionCloseImp::where('model_unique_code', $model_unique_code)
            ->where('is_agree_substitution', '是')
            ->selectRaw('cpn_id, concat(cpn_specification_model,cpn_manufacturer) as spe_manufacturer, concat(kb_substitution_model,kb_substitution_mfr) as replace_spe_manufacturer')
            ->get()
            ->toArray();
        foreach ($source as $key => $val)
        {
            if (!in_array($val['spe_manufacturer'], $importData) && in_array($val['replace_spe_manufacturer'], $domesticData))
            {
                $data[] = [
                    'main_task_id' => $main_task_id,
                    'list_id'      => '',
                    'cpn_type'     => '',
                    'cpn_id'       => $val['cpn_id'],
                    'rule_id'      => 2,
                    'messages'     => '已落实',
                ];
            } elseif (in_array($val['spe_manufacturer'], $importData) && !in_array($val['replace_spe_manufacturer'], $domesticData))
            {
                $data[] = [
                    'main_task_id' => $main_task_id,
                    'list_id'      => '',
                    'cpn_type'     => '',
                    'cpn_id'       => $val['cpn_id'],
                    'rule_id'      => 2,
                    'messages'     => '未落实',
                ];
            } else
            {
                $data[] = [
                    'main_task_id' => $main_task_id,
                    'list_id'      => '',
                    'cpn_type'     => '',
                    'cpn_id'       => $val['cpn_id'],
                    'rule_id'      => 2,
                    'messages'     => '数据异常',
                ];
            }
        }

        //更新完成任务数
        MainTask::where('id', $main_task_id)->increment('auxiliary_checked_num');

        return $data;
    }

    /**
     * 研制单位反馈存在伪空包现象的数据
     * @param $main_task_id
     * @param $model_unique_code
     * @param $reviewList1
     * @return array
     */
    public function implementFake($main_task_id = null, $model_unique_code = null, $reviewList1 = [])
    {
        if (empty($main_task_id) || empty($model_unique_code) || empty($reviewList1))
            return [];


        $data         = [];//审查结果数据
        $domesticData = array_column($reviewList1, 'spe_manufacturer');//国产清单数据
        $source       = ProReviewOpinionCloseDom::where('model_unique_code', $model_unique_code)
            ->where('is_agree_empty_package', '是')
            ->selectRaw('cpn_id, concat(cpn_specification_model,cpn_manufacturer) as spe_manufacturer')
            ->get()
            ->toArray();
        foreach ($source as $key => $val)
        {
            if (!in_array($val['spe_manufacturer'], $domesticData))
            {
                $data[] = [
                    'main_task_id' => $main_task_id,
                    'list_id'      => '',
                    'cpn_type'     => '',
                    'cpn_id'       => $val['cpn_id'],
                    'rule_id'      => 3,
                    'messages'     => '已整改剔除',
                ];
            } else
            {
                $data[] = [
                    'main_task_id' => $main_task_id,
                    'list_id'      => '',
                    'cpn_type'     => '',
                    'cpn_id'       => $val['cpn_id'],
                    'rule_id'      => 3,
                    'messages'     => '未剔除',
                ];
            }
        }

        //更新完成任务数
        MainTask::where('id', $main_task_id)->increment('auxiliary_checked_num');

        return $data;
    }

    /**
     * 伪国产化、空心国产化、包装国产化知识库
     * @param $main_task_id
     * @param $reviewList1
     * @return array
     */
    public function fakeEmpty($main_task_id = null, $reviewList1 = [])
    {
        if (empty($main_task_id) || empty($reviewList1))
            return [];


        $data       = [];//审查结果数据
        $source     = ProReviewOpinionCloseDom::where('is_agree_empty_package', '是')
            ->selectRaw('id, concat(cpn_specification_model,cpn_manufacturer) as spe_manufacturer')
            ->get()
            ->toArray();
        $sourceData = array_column($source, 'spe_manufacturer');
        foreach ($reviewList1 as $key => $val)
        {
            if (in_array($val['spe_manufacturer'], $sourceData))
            {
                $data[] = [
                    'main_task_id' => $main_task_id,
                    'list_id'      => $val['list_id'],
                    'cpn_type'     => 0,
                    'cpn_id'       => $val['id'],
                    'rule_id'      => 4,
                    'messages'     => '是',
                ];
            } else
            {
                $data[] = [
                    'main_task_id' => $main_task_id,
                    'list_id'      => $val['list_id'],
                    'cpn_type'     => 0,
                    'cpn_id'       => $val['id'],
                    'rule_id'      => 4,
                    'messages'     => '否',
                ];
            }
        }

        //更新完成任务数
        MainTask::where('id', $main_task_id)->increment('auxiliary_checked_num');

        return $data;
    }

    /**
     * 历史伪空包现象专家审查数据
     * @param $main_task_id
     * @param $reviewList1
     * @return array
     */
    public function historyFakeEmpty($main_task_id = null, $reviewList1 = [])
    {
        if (empty($main_task_id) || empty($reviewList1))
            return [];

        $data       = [];//审查结果数据
        $source     = ProReviewOpinionDomestic::where('kb_empty_whether', '是')
            ->selectRaw('expert_name, concat(cpn_specification_model,cpn_manufacturer) as spe_manufacturer')
            ->groupBy('expert_name', 'spe_manufacturer')
            ->get()
            ->toArray();
        $sourceData = array_column($source, 'spe_manufacturer');
        foreach ($reviewList1 as $key => $val)
        {
            //存在数据源中
            if ($ks = array_keys($sourceData, $val['spe_manufacturer']))
            {
                $data[] = [
                    'main_task_id' => $main_task_id,
                    'list_id'      => $val['list_id'],
                    'cpn_type'     => 0,
                    'cpn_id'       => $val['id'],
                    'rule_id'      => 5,
                    'messages'     => count($ks) . '个专家识别有问题',
                ];
            }
        }

        //更新完成任务数
        MainTask::where('id', $main_task_id)->increment('auxiliary_checked_num');

        return $data;
    }

    /**
     * 历史满足度专家审查数据
     * @param $main_task_id
     * @param $reviewList1
     * @return array
     */
    public function historySatisfy($main_task_id = null, $reviewList1 = [])
    {
        if (empty($main_task_id) || empty($reviewList1))
            return [];

        $data       = [];//审查结果数据
        $source     = ProReviewOpinionDomestic::where('kb_satisfaction_whether', '是')
            ->selectRaw('expert_name, concat(cpn_specification_model,cpn_manufacturer) as spe_manufacturer')
            ->groupBy('expert_name', 'spe_manufacturer')
            ->get()
            ->toArray();
        $sourceData = array_column($source, 'spe_manufacturer');
        foreach ($reviewList1 as $key => $val)
        {
            //存在数据源中
            if ($ks = array_keys($sourceData, $val['spe_manufacturer']))
            {
                $data[] = [
                    'main_task_id' => $main_task_id,
                    'list_id'      => $val['list_id'],
                    'cpn_type'     => 0,
                    'cpn_id'       => $val['id'],
                    'rule_id'      => 6,
                    'messages'     => count($ks) . '个专家识别有风险',
                ];
            }
        }

        //更新完成任务数
        MainTask::where('id', $main_task_id)->increment('auxiliary_checked_num');

        return $data;
    }

    /**
     * 对进口产品（技术）YCD数据（电子元器件类）
     * @param $main_task_id
     * @param $reviewList2
     * @return array
     */
    public function ycd($main_task_id = null, $reviewList2 = [])
    {
        if (empty($main_task_id) || empty($reviewList2))
            return [];

        $data       = [];//审查结果数据
        $source     = DataDepend::select('id', 'cpn_specification_model', 'depend_level')->get()->toArray();
        $sourceData = array_column($source, 'cpn_specification_model');
        foreach ($reviewList2 as $key => $val)
        {
            //存在数据源中
            if (false !== ($k = array_search($val['cpn_specification_model'], $sourceData)))
            {
                $data[] = [
                    'id'         => $val['id'],
                    'dependence' => $source[$k]['depend_level'],
                ];
            }
        }

        //更新数据
        $CpnImport = new CpnImport();
        $CpnImport->updateBatch($data);

        //更新完成任务数
        MainTask::where('id', $main_task_id)->increment('auxiliary_checked_num');

        return true;
    }

    /**
     * COTS类器件质量等级字典
     * @param $main_task_id
     * @param $reviewList2
     * @return array
     */
    public function cots($main_task_id = null, $reviewList2 = [])
    {
        if (empty($main_task_id) || empty($reviewList2))
            return [];

        $data = [];//审查结果数据
        foreach ($reviewList2 as $key => $val)
        {
            //存在数据源中
            $k = array_search($val['cpn_quality'], $this->CotsData);
            if ($k !== false)
            {
                $data[] = [
                    'id'      => $val['id'],
                    'is_cots' => 1,
                ];
            }
        }

        //更新数据
        $CpnImport = new CpnImport();
        $CpnImport->updateBatch($data);

        //更新完成任务数
        MainTask::where('id', $main_task_id)->increment('auxiliary_checked_num');

        return true;
    }

    /**
     * 进口清单跨阶段比对审查
     * @param null  $main_task_id
     * @param array $reviewList2
     * @param null  $contrast_main_task_id
     * @return array
     */
    public function crossStage($main_task_id = null, $reviewList2 = [], $contrast_main_task_id = null)
    {
        if (empty($main_task_id) || empty($reviewList2) || empty($contrast_main_task_id))
            return [];

        $data = [];//审查结果数据
        //获取对比版本清单信息
        $cmpFilesIds   = CpnFiles::where('main_task_id', $contrast_main_task_id)
            ->where('type', '进口')
            ->pluck('id')
            ->toArray();
        $cmpImport     = CpnImport::whereIn('list_id', $cmpFilesIds)
            ->selectRaw('id, list_id, cpn_specification_model, cpn_manufacturer, concat(cpn_specification_model,cpn_manufacturer) as spe_manufacturer')
            ->get()
            ->toArray();
        $cmpImportData = array_column($cmpImport, 'spe_manufacturer');

        //当前版本数据
        $importData = array_column($reviewList2, 'spe_manufacturer');

        //超出
        foreach ($reviewList2 as $key => $val)
        {
            if (!in_array($val['spe_manufacturer'], $cmpImportData))
            {
                $data[] = [
                    'main_task_id' => $main_task_id,
                    'list_id'      => $val['list_id'],
                    'cpn_type'     => 1,
                    'cpn_id'       => $val['id'],
                    'rule_id'      => 15,
                    'messages'     => '超出',
                ];
            }
        }

        //更新完成任务数
        MainTask::where('id', $main_task_id)->increment('auxiliary_checked_num');

        //统计减少的数据
        foreach ($cmpImport as $key => $val)
        {
            if (!in_array($val['spe_manufacturer'], $importData))
            {
                $data[] = [
                    'main_task_id' => $main_task_id,
                    'list_id'      => $val['list_id'],
                    'cpn_type'     => 1,
                    'cpn_id'       => $val['id'],
                    'rule_id'      => 15,
                    'messages'     => '减少',
                ];
            }
        }

        return $data;
    }

    /**
     * 进口清单国产化替代
     * @param null  $main_task_id
     * @param array $listIds
     * @return array
     * @throws \Exception
     */
    public function domesticReplace($main_task_id = null, $listIds = [])
    {
        if (empty($main_task_id) || empty($listIds))
            return [];

        $newData     = [];//新数据
        $replaceData = [];//审查结果数据
        $chunkSize   = 200;

        $reviewList = CpnImport::whereIn('list_id', $listIds)
            ->where('is_repeat', 0)//过滤重复项
            //            ->where('result_grc', '>', 1)//合规性审查
            ->selectRaw('id, list_id, cpn_specification_model, cpn_ref_price, cpn_period, result_grc, yield_is_core_important, yield_safe_color')
            //            ->groupBy('cpn_specification_model')
            ->orderBy('id', 'asc')
            ->get()
            ->toArray(); //进口数据

        //ES初始化
        $model = new Elasticsearch();
        $model->setRulesByTask($main_task_id);

        $chunks = collect($reviewList)->chunk($chunkSize);
        foreach ($chunks as $block)
        {
            $rows = $block->toArray();
            $csms = array_column($rows, 'cpn_specification_model');
            //国产军用电子元器件手册
            $list = AuxiliaryDataDomesticMilitary::whereIn('cpn_specification_model', $csms)
                ->select('cpn_specification_model', 'replace_cpn_specification_model', 'replace_cpn_manufacturer', 'replace_cpn_quality', 'replace_type')
                ->get()
                ->toArray();
            //赛思库®电子元器件国产化替代数据
            $list2 = AuxiliaryDataDomesticCiss::whereIn('cpn_specification_model', $csms)
                ->select('id', 'cpn_specification_model', 'replace_cpn_specification_model', 'replace_cpn_manufacturer', 'replace_cpn_quality', 'replace_type', 'replace_product_state', 'hash')
                ->get()
                ->toArray();
            //型号规格集合
            $contrast  = array_column($list, 'cpn_specification_model');
            $contrast2 = array_column($list2, 'cpn_specification_model');
            foreach ($rows as $key => $val)
            {
                //是否在国产军用电子元器件手册
                if ($ks = array_keys($contrast, $val['cpn_specification_model']))
                {
                    $flag = false;
                    foreach ($ks as $k)
                    {
                        //替代类型是否为“原位替代”
                        if ($list[$k]['replace_type'] == '原位替代')
                        {
                            $flag = true;
                        }

                        //可替代产品应用经历
                        $experience = $this->getExperience($list[$k]['replace_cpn_specification_model']);
                        //替代信息
                        $replaceData[] = [
                            'main_task_id'                    => $main_task_id,
                            'list_id'                         => $val['list_id'],
                            'cpn_id'                          => $val['id'],
                            'source'                          => 1,
                            'replace_cpn_specification_model' => $list[$k]['replace_cpn_specification_model'],
                            'replace_cpn_manufacturer'        => $list[$k]['replace_cpn_manufacturer'],
                            'replace_cpn_quality'             => $list[$k]['replace_cpn_quality'],
                            'replace_type'                    => '',
                            'replace_product_state'           => '',
                            'experience'                      => $experience,
                        ];
                    }

                    $rows[$key]['dependence'] = $flag ? '三级' : '二级';//依存度
                } elseif ($ks = array_keys($contrast2, $val['cpn_specification_model']))
                { //是否在赛思库®电子元器件国产化替代数据
                    $flag       = false;
                    $replaceIds = [];//替代信息的ids
                    foreach ($ks as $k)
                    {
                        //替代类型是否为“原位替代”
                        if ($list2[$k]['replace_type'] == '原位替代')
                        {
                            $flag = true;
                        }

                        //可替代产品应用经历
                        $experience = $this->getExperience($list2[$k]['replace_cpn_specification_model']);

                        //替代信息
                        $replaceData[] = [
                            'main_task_id'                    => $main_task_id,
                            'list_id'                         => $val['list_id'],
                            'cpn_id'                          => $val['id'],
                            'source'                          => 2,
                            'replace_cpn_specification_model' => $list2[$k]['replace_cpn_specification_model'],
                            'replace_cpn_manufacturer'        => $list2[$k]['replace_cpn_manufacturer'],
                            'replace_cpn_quality'             => $list2[$k]['replace_cpn_quality'],
                            'replace_type'                    => $list2[$k]['replace_type'],
                            'replace_product_state'           => $list2[$k]['replace_product_state'],
                            'experience'                      => $experience,
                        ];

                        $replaceIds[] = $list2[$k]['hash'];
                    }

                    $rows[$key]['dependence'] = $flag ? '三级' : '二级';//依存度
                    //当判定ycd为三级时，添加替代产品的价格、货期
                    if ($flag && !empty($replaceIds))
                    {
                        if ($replaceInfo = $this->getReplaceInfo($replaceIds))
                        {
                            $rows[$key]['replace_price']    = $replaceInfo['price'];//价格
                            $rows[$key]['replace_delivery'] = $replaceInfo['delivery'];//货期
                        }
                    }
                } else
                {//是否在国产军用电子元器件手册中有同类产品
                    $catCodes = CpnImport::whereIn('list_id', $listIds)
                        ->where('cpn_specification_model', $val['cpn_specification_model'])
                        ->groupBy('cpn_category_code')
                        ->pluck('cpn_category_code')
                        ->toArray(); //分类代码集合
                    if ($res = AuxiliaryDataDomesticMilitary::whereIn('category_code', $catCodes)->first())
                    {
                        $rows[$key]['dependence'] = '二级';//依存度
                    } else
                    {//数据源参数计算是否有相似产品
                        $result = $model->getSimilarComponents($val['cpn_specification_model']);
                        if (!empty($result['replace_cpns']))
                        {
                            $rows[$key]['dependence'] = '二级';//依存度
                            foreach ($result['replace_cpns'] as $k => $v)
                            {
                                //可替代产品应用经历
                                $experience    = $this->getExperience($v['part_number']);
                                $replaceData[] = [
                                    'main_task_id'                    => $main_task_id,
                                    'list_id'                         => $val['list_id'],
                                    'cpn_id'                          => $val['id'],
                                    'source'                          => 2,
                                    'replace_cpn_specification_model' => $v['part_number'],
                                    'replace_cpn_manufacturer'        => $v['part_manufacturer'],
                                    'replace_cpn_quality'             => $v['quality_level'],
                                    'replace_type'                    => '',
                                    'replace_product_state'           => '',
                                    'experience'                      => $experience,
                                ];
                            }
                        } else
                        {
                            //必要性
                            $necessityArr = CpnImport::whereIn('list_id', $listIds)
                                ->where('cpn_specification_model', $val['cpn_specification_model'])
                                ->pluck('yield_necessity')
                                ->toArray();

                            $rows[$key]['dependence'] = ($val['yield_is_core_important'] == 1) &&
                            ((count($necessityArr) == 1) && ($necessityArr[0] == 1))
                                ? '一级' : '二级'; //依存度
                        }
                    }
                }
            }

            $newData = array_merge($newData, $rows);
        }
        //        dd($newData, $replaceData);

        return [
            'newData'     => $newData,
            'replaceData' => $replaceData,
        ];
    }

    /**
     * 历史安全性审查专家审查数据
     * @param $main_task_id
     * @param $reviewList2
     * @return array
     */
    public function doColor($main_task_id = null, $reviewList2 = [])
    {
        if (empty($main_task_id) || empty($reviewList2))
            return [];

        $data = [];//审查结果数据
        //历史安全性知识库
        $sourceSafe     = ProReviewOpinionImport::where('kb_safety_whether', '是')
            ->selectRaw('expert_name, cpn_specification_model')
            ->groupBy('expert_name', 'cpn_specification_model')
            ->get()
            ->toArray();
        $sourceSafeData = array_column($sourceSafe, 'cpn_specification_model');

        //历史可保障性知识库
        $sourceEnsure     = ProReviewOpinionImport::where('kb_insurability_whether', '是')
            ->select('expert_name', 'cpn_specification_model')
            ->groupBy('expert_name', 'cpn_specification_model')
            ->get()
            ->toArray();
        $sourceEnsureData = array_column($sourceEnsure, 'cpn_specification_model');

        //开始比对
        foreach ($reviewList2 as $key => $val)
        {
            //比对进口清单“规格型号、生产厂商”是否在“国产手册颜色等级库”中
            $cpn_manufacturer = $val['cpn_manufacturer'];
            $sourceData       = DataColor::where('part_number', $val['cpn_specification_model'])
                ->where(function ($query) use ($cpn_manufacturer) {
                    $query->where('manufacturer', $cpn_manufacturer)
                        ->orWhere('manufacturer_simplify', $cpn_manufacturer);
                })
                ->select('color')
                ->first(); //颜色等级数据源
            if (!empty($sourceData))
            {//存在数据源中
                //获取最高级别的颜色等级
                if (false !== strpos($val['safe_color'], '、'))
                {
                    $arr    = explode('、', $val['safe_color']);
                    $colors = $this->ColorGrade;
                    foreach ($colors as $k => $v)
                    {
                        if (!in_array($v, $arr))
                        {
                            unset($colors[$k]);
                        }
                    }
                    krsort($colors);
                    $val['color'] = current($colors);
                } else
                {
                    $val['color'] = $val['safe_color'];
                }

                //比对进口清单中“安全颜色等级”与数据源中是否一致
                if ($val['color'] == $sourceData['color'])
                {//一致
                    if ($val['color'] == '红色' || $val['color'] == '紫色')
                    {
                        //将“确认或可能存在安全风险”标记为“确认”
                        $data[] = [
                            'main_task_id'  => $main_task_id,
                            'list_id'       => $val['list_id'],
                            'cpn_type'      => 1,
                            'cpn_id'        => $val['id'],
                            'rule_id'       => 13,
                            'risk'          => '确认',
                            'messages'      => '',
                            'color'         => $val['safe_color'],
                            'suggest_color' => '',
                        ];
                    } elseif ($val['color'] == '橙色' || $val['color'] == '黄色')
                    {
                        //将“确认或可能存在可保障风险”标记为“确认”
                        $data[] = [
                            'main_task_id'  => $main_task_id,
                            'list_id'       => $val['list_id'],
                            'cpn_type'      => 1,
                            'cpn_id'        => $val['id'],
                            'rule_id'       => 14,
                            'risk'          => '确认',
                            'messages'      => '',
                            'color'         => $val['safe_color'],
                            'suggest_color' => '',
                        ];
                    }
                } else
                {//不一致
                    //进口清单中“安全颜色等级”是否都不高于数据源中的“安全颜色等级”
                    if ($res = $this->isLteColor($val['color'], $sourceData['color']))
                    {//不高于
                        if ($sourceData['color'] == '红色' || $sourceData['color'] == '紫色')
                        {
                            //将“确认或可能存在安全风险”标记为“确认”
                            $data[] = [
                                'main_task_id'  => $main_task_id,
                                'list_id'       => $val['list_id'],
                                'cpn_type'      => 1,
                                'cpn_id'        => $val['id'],
                                'rule_id'       => 13,
                                'risk'          => '确认',
                                'messages'      => '',
                                'color'         => $val['safe_color'],
                                'suggest_color' => $sourceData['color'],
                            ];
                        } elseif ($sourceData['color'] == '橙色' || $sourceData['color'] == '黄色')
                        {
                            //将“确认或可能存在可保障风险”标记为“确认”
                            $data[] = [
                                'main_task_id'  => $main_task_id,
                                'list_id'       => $val['list_id'],
                                'cpn_type'      => 1,
                                'cpn_id'        => $val['id'],
                                'rule_id'       => 14,
                                'risk'          => '确认',
                                'messages'      => '',
                                'color'         => $val['safe_color'],
                                'suggest_color' => $sourceData['color'],
                            ];
                        }
                    } else
                    {//高于
                        if ($val['color'] == '紫色')
                        {
                            //是否在历史安全性专家审查数据中
                            if ($ks = array_keys($sourceSafeData, $val['cpn_specification_model']))
                            {
                                $data[] = [
                                    'main_task_id'  => $main_task_id,
                                    'list_id'       => $val['list_id'],
                                    'cpn_type'      => 1,
                                    'cpn_id'        => $val['id'],
                                    'rule_id'       => 13,
                                    'risk'          => '可能',
                                    'messages'      => count($ks) . '个专家识别有风险',
                                    'color'         => $val['safe_color'],
                                    'suggest_color' => '',
                                ];
                            } else
                            {
                                $data[] = [
                                    'main_task_id'  => $main_task_id,
                                    'list_id'       => $val['list_id'],
                                    'cpn_type'      => 1,
                                    'cpn_id'        => $val['id'],
                                    'rule_id'       => 14,
                                    'risk'          => '确认',
                                    'messages'      => '',
                                    'color'         => $val['safe_color'],
                                    'suggest_color' => '',
                                ];
                            }
                        } elseif ($val['color'] == '橙色' && $sourceData['color'] == '黄色')
                        {
                            $data[] = [
                                'main_task_id'  => $main_task_id,
                                'list_id'       => $val['list_id'],
                                'cpn_type'      => 1,
                                'cpn_id'        => $val['id'],
                                'rule_id'       => 14,
                                'risk'          => '确认',
                                'messages'      => '',
                                'color'         => $val['safe_color'],
                                'suggest_color' => '',
                            ];
                        } elseif (($val['color'] == '橙色' || $val['color'] == '黄色') && $sourceData['color'] == '绿色')
                        {
                            //是否在历史可保障性专家审查数据中
                            if ($ks = array_keys($sourceEnsureData, $val['cpn_specification_model']))
                            {
                                $data[] = [
                                    'main_task_id'  => $main_task_id,
                                    'list_id'       => $val['list_id'],
                                    'cpn_type'      => 1,
                                    'cpn_id'        => $val['id'],
                                    'rule_id'       => 14,
                                    'risk'          => '可能',
                                    'messages'      => count($ks) . '个专家识别有风险',
                                    'color'         => $val['safe_color'],
                                    'suggest_color' => '',
                                ];
                            }
                        }
                    }
                }
            } else
            {//不存在数据源中
                //是否在历史安全性专家审查数据中
                if ($ks = array_keys($sourceSafeData, $val['cpn_specification_model']))
                {
                    $data[] = [
                        'main_task_id'  => $main_task_id,
                        'list_id'       => $val['list_id'],
                        'cpn_type'      => 1,
                        'cpn_id'        => $val['id'],
                        'rule_id'       => 13,
                        'risk'          => '可能',
                        'messages'      => count($ks) . '个专家识别有风险',
                        'color'         => $val['safe_color'],
                        'suggest_color' => '',
                    ];
                }

                //是否在历史可保障性专家审查数据中
                if ($ks = array_keys($sourceEnsureData, $val['cpn_specification_model']))
                {
                    $data[] = [
                        'main_task_id'  => $main_task_id,
                        'list_id'       => $val['list_id'],
                        'cpn_type'      => 1,
                        'cpn_id'        => $val['id'],
                        'rule_id'       => 14,
                        'risk'          => '可能',
                        'messages'      => count($ks) . '个专家识别有风险',
                        'color'         => $val['safe_color'],
                        'suggest_color' => '',
                    ];
                }
            }
        }

        //更新完成任务数
        MainTask::where('id', $main_task_id)->increment('auxiliary_checked_num', 2);

        return $data;
    }

    /**
     * 判断颜色等级是否不高于数据源
     * @param $color
     * @param $sourceColor
     * @return bool
     */
    public function isLteColor($color, $sourceColor)
    {
        if (empty($color) || empty($sourceColor))
            return false;

        $k1 = array_search($color, $this->ColorGrade);
        $k2 = array_search($sourceColor, $this->ColorGrade);

        return ($k1 <= $k2) ? true : false;
    }


    /**
     * 根据某 多个字段去重二位数组
     * @param array $arr
     * @param       $filter
     * @return array
     */
    function array_unique_fb($arr = [], $filter)
    {
        $res = [];
        foreach ($arr as $key => $value)
        {
            $newkey = "";
            if (is_array($filter))
            {
                foreach ($filter as $fv)
                {
                    $newkey .= $value[$fv];
                }
            } else
            {
                $newkey = $value[$filter];
            }
            foreach ($value as $vk => $va)
            {
                if (isset($res[$newkey]))
                {
                    $res[$newkey][$vk] = $va;
                } else
                {
                    $res[$newkey][$vk] = $va;
                }
            }
        }
        return $res;
    }

    /**
     * 获取替代信息
     * @param $replaceIds
     * @return array
     */
    public function getReplaceInfo($replaceIds)
    {
        $data = AuxiliaryDataDomesticCissHistory::whereIn('hash', $replaceIds)->select('price', 'delivery')->get();
        if ($data->isEmpty())
            return [];

        $data = $data->toArray();
        //格式化数据
        $formatVal   = function ($v) {
            if (strpos($v, '_') !== false)
            {
                $arr = explode('_', $v);
                return $arr[1];
            }
            return $v;
        };
        $priceArr    = array_map($formatVal, array_column($data, 'price'));
        $deliveryArr = array_map($formatVal, array_column($data, 'delivery'));

        return [
            'price'    => max($priceArr),
            'delivery' => max($deliveryArr),
        ];
    }

    /**
     * 获取清单与权重对应map
     * @param $main_task_id
     * @return array
     */
    public function getListWeightMap($main_task_id)
    {
        $structures = ModelStructure::where('main_task_id', $main_task_id)->get()->toArray();

        // 取出所有叶子结点
        $parent_ids          = array_column($structures, 'parent_id');
        $structureCollection = Collection::make($structures);

        $leafNodes = $structureCollection->filter(function ($item) use ($parent_ids) {
            return !in_array($item['id'], $parent_ids);
        })->all();

        $structure_ids = array_values(array_column($leafNodes, 'id'));
        $lists         = CpnFiles::whereIn('model_id', $structure_ids)
            ->select('id', 'model_id')
            ->get()
            ->toArray();

        // 根据code拆分，查询各层级结点的权重并相乘
        foreach ($leafNodes as $key => $node)
        {
            // 父结点编码
            $deep = strlen($node['code']) / 3;
            // 当前路径权重
            $weight = 1;
            for ($i = 1; $i <= $deep; $i++)
            {
                // 截取编码
                $code = substr($node['code'], 0, $i * 3);
                // 取当前结点
                $tempNode = $structureCollection->where('code', $code)->first();
                // 计算权重
                $weight = $weight * $tempNode['number'];
            }
            $pathWeight[$node['id']] = $weight;
        }

        // 清单id与权重对应表
        $listWeightMap = [];
        foreach ($lists as $key => $list)
        {
            $listWeightMap[$list['id']]['weight']   = $pathWeight[$list['model_id']];
            $listWeightMap[$list['id']]['model_id'] = $list['model_id'];
        }

        return $listWeightMap;
    }

    /**
     * 获取伪空包推荐数据
     * @param       $main_task_id
     * @param array $data
     * @return array
     */
    public function getRecommendData($main_task_id = '', $data = [], $type = 1)
    {
        if (empty($data) || empty($main_task_id))
            return [];

        arsort($data);
        $top5 = array_slice($data, 0, ceil(count($data) * 0.05));

        $recommendData = [];
        foreach ($top5 as $k => $v)
        {
            $arr             = explode('@', $k);
            $recommendData[] = [
                'main_task_id'            => $main_task_id,
                'cpn_specification_model' => $arr[0],
                'cpn_manufacturer'        => $arr[1],
                'type'                    => $type
            ];
        }

        return $recommendData;
    }

    /**
     * 获取意见判定结果
     * @param array $newData
     * @return array
     */
    public function getOpinionData($newData = [])
    {
        if (empty($newData))
            return [];

        $data = [];
        foreach ($newData as $key => $val)
        {
            //是否核关高
            $isCore = $val['yield_is_core_important'];
            //安全颜色等级
            $safeColor = $val['yield_safe_color'];
            //价格
            if (strpos($val['cpn_ref_price'], '_') !== false)
            {
                $arr   = explode('_', $val['cpn_ref_price']);
                $price = $arr[1];
            } else
            {
                $price = $val['cpn_ref_price'];
            }
            //货期
            if (strpos($val['cpn_period'], '_') !== false)
            {
                $arr    = explode('_', $val['cpn_period']);
                $period = $arr[1];
            } else
            {
                $period = $val['cpn_period'];
            }

            //让步接收
            if ($val['result_grc'] == 2)
            {
                $result_pc = 4;//进入用研结合审查
            } else
            {//非让步接收
                if ($safeColor === '红')
                {
                    $newData[$key]['result_pc'] = 3;//不允许选用
                } else
                {
                    //依存度一级
                    if ($val['dependence'] == '一级')
                    {
                        $result_pc = 1;//纳入进口电子元器件清单
                    } elseif ($val['dependence'] === '二级')
                    {//依存度二级
                        if (($safeColor === '绿') && ($isCore === 0))
                        {
                            $result_pc = 1;//纳入进口电子元器件清单
                        } else
                        {
                            $result_pc = 4;//进入用研结合审查
                        }
                    } else
                    {//依存度三级
                        if (($safeColor === '绿') && ($isCore === 0))
                        {
                            $result_pc = 1;//纳入进口电子元器件清单
                        } elseif ($isCore === 1)
                        {
                            if (($price < 1000) && (isset($val['replace_price']) && ($val['replace_price'] > $price * 3)))
                            {
                                $result_pc = 1;//纳入进口电子元器件清单
                            } elseif (($price > 1000) && (isset($val['replace_price']) && ($val['replace_price'] > $price * 2)))
                            {
                                $result_pc = 1;//纳入进口电子元器件清单
                            } elseif (isset($val['replace_delivery']) && ($val['replace_delivery'] > ($period + 12)))
                            {
                                $result_pc = 1;//纳入进口电子元器件清单
                            } else
                            {
                                $result_pc = 2;//纳入国产电子元器件清单
                            }
                        } else
                        {
                            $result_pc = 2;//纳入国产电子元器件清单
                        }
                    }
                }
            }

            $data[] = [
                'id'         => $val['id'],
                'dependence' => $val['dependence'],
                'result_pc'  => $result_pc
            ];
        }

        return $data;
    }

    /**
     * 获取伪空包推荐结果
     * @param string $main_task_id
     * @param array  $listIds
     * @return array
     */
    public function getFakeData($main_task_id = '', $listIds = [])
    {
        if (empty($main_task_id) || empty($listIds))
            return [];

        $listWeightMap = $this->getListWeightMap($main_task_id);
        $reviewList    = CpnDomestic::whereIn('list_id', $listIds)
            ->where('yield_is_core_important', 1)//核关高
            ->where('is_repeat', 0)//过滤重复项
            //            ->where('result_grc', '>', 1)//合规性审查
            ->selectRaw('id, list_id, equip_use_number, cpn_ref_price, concat_ws("@", cpn_specification_model,cpn_manufacturer) as spem, yield_control_level')
            ->orderBy('id', 'asc')
            ->get()
            ->toArray();//国产数据
        if (empty($reviewList))
            return [];

        $amountData   = [];//使用数量
        $amountDataC  = [];//自主可控等级C使用数量
        $amountDataD  = [];//自主可控等级D使用数量
        $amountDataE  = [];//自主可控等级E使用数量
        $positionData = [];//装机位置
        $priceDataA   = [];//自主可控等级A单价
        $priceDataB   = [];//自主可控等级B单价
        $priceDataC   = [];//自主可控等级C单价
        foreach ($reviewList as $key => $val)
        {
            //自主可控等级
            $level = $val['yield_control_level'];
            //价格
            if (strpos($val['cpn_ref_price'], '_') !== false)
            {
                $arr   = explode('_', $val['cpn_ref_price']);
                $price = $arr[1];
            } else
            {
                $price = $val['cpn_ref_price'];
            }

            $weight = @$listWeightMap[$val['list_id']]['weight'];
            $amount = $val['equip_use_number'] * $weight;//单节点使用数量
            @$amountData[$val['spem']] += $amount;
            @$positionData[$val['spem']] += 1;
            if ($level == 'A')
            {
                $priceDataA[$val['spem']] = $price;
            } elseif ($level == 'B')
            {
                $priceDataB[$val['spem']] = $price;
            } elseif ($level == 'C')
            {
                @$amountDataC[$val['spem']] += $amount;
                $priceDataC[$val['spem']] = $price;
            } elseif ($level == 'D')
            {
                @$amountDataD[$val['spem']] += $amount;
            } elseif ($level == 'E')
            {
                @$amountDataE[$val['spem']] += $amount;
            }
        }

        $recommendData = $this->getRecommendData($main_task_id, $amountData, 1);
        $recommendData = array_merge($recommendData, $this->getRecommendData($main_task_id, $positionData, 2));
        $recommendData = array_merge($recommendData, $this->getRecommendData($main_task_id, $amountDataC, 3));
        $recommendData = array_merge($recommendData, $this->getRecommendData($main_task_id, $amountDataD, 4));
        $recommendData = array_merge($recommendData, $this->getRecommendData($main_task_id, $amountDataE, 5));
        $recommendData = array_merge($recommendData, $this->getRecommendData($main_task_id, $priceDataA, 6));
        $recommendData = array_merge($recommendData, $this->getRecommendData($main_task_id, $priceDataB, 7));
        $recommendData = array_merge($recommendData, $this->getRecommendData($main_task_id, $priceDataC, 8));

        return $recommendData;
    }

}