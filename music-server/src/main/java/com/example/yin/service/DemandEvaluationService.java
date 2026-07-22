package com.example.yin.service;

import com.baomidou.mybatisplus.extension.service.IService;
import com.example.yin.common.R;
import com.example.yin.model.domain.DemandEvaluation;

import java.util.List;

/**
 * 针对表【demand_evaluation】的数据库操作Service
 */
public interface DemandEvaluationService extends IService<DemandEvaluation> {

    /**
     * 获取所有需求
     */
    R getAllDemand();

    /**
     * 根据ID获取需求详情
     */
    R getDemandOfId(Long id);

    /**
     * 添加需求
     */
    R addDemand(DemandEvaluation demand);

    /**
     * 更新需求信息
     */
    R updateDemand(DemandEvaluation demand);

    /**
     * 删除需求
     */
    R deleteDemand(Long id);
}
