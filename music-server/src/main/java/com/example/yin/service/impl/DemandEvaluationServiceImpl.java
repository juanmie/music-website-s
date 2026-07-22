package com.example.yin.service.impl;

import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import com.example.yin.common.R;
import com.example.yin.mapper.DemandEvaluationMapper;
import com.example.yin.model.domain.DemandEvaluation;
import com.example.yin.service.DemandEvaluationService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

/**
 * 针对表【demand_evaluation】的数据库操作Service实现
 */
@Service
public class DemandEvaluationServiceImpl extends ServiceImpl<DemandEvaluationMapper, DemandEvaluation>
        implements DemandEvaluationService {

    @Autowired
    private DemandEvaluationMapper demandEvaluationMapper;

    @Override
    public R getAllDemand() {
        List<DemandEvaluation> list = demandEvaluationMapper.selectList(null);
        return R.success("成功获取所有需求", list);
    }

    @Override
    public R getDemandOfId(Long id) {
        DemandEvaluation demand = demandEvaluationMapper.selectById(id);
        return R.success("成功获取需求详情", demand);
    }

    @Override
    public R addDemand(DemandEvaluation demand) {
        int result = demandEvaluationMapper.insert(demand);
        if (result > 0) {
            return R.success("添加需求成功");
        }
        return R.error("添加需求失败");
    }

    @Override
    public R updateDemand(DemandEvaluation demand) {
        int result = demandEvaluationMapper.updateById(demand);
        if (result > 0) {
            return R.success("更新需求成功");
        }
        return R.error("更新需求失败");
    }

    @Override
    public R deleteDemand(Long id) {
        int result = demandEvaluationMapper.deleteById(id);
        if (result > 0) {
            return R.success("删除需求成功");
        }
        return R.error("删除需求失败");
    }
}
