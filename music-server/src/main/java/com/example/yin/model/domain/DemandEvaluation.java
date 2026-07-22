package com.example.yin.model.domain;

import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableField;
import com.baomidou.mybatisplus.annotation.TableId;
import com.baomidou.mybatisplus.annotation.TableName;
import lombok.Data;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.Date;

/**
 * 需求评估导入表
 * @TableName demand_evaluation
 */
@Data
@TableName("demand_evaluation")
public class DemandEvaluation implements Serializable {

    @TableId(type = IdType.AUTO)
    private Long id;

    /** 序号 */
    private Integer seqNo;

    /** 程序代号 */
    private String programCode;

    /** 程序名称 */
    private String programName;

    /** 计费时数-客户 */
    private BigDecimal billingHoursCustomer;

    /** 软代派工时数 */
    private BigDecimal dispatchHours;

    /** 系统代号 */
    private String systemCode;

    /** 规格说明 */
    private String specDesc;

    /** 备注 */
    private String remark;

    /** 客制类型 */
    private String customType;

    /** 难度等级 */
    private String difficultyLevel;

    /** 开发分类 */
    private String devCategory;

    /** 集成产品编号 */
    private String integrateProductCode;

    /** 急单 */
    private String urgentFlag;

    /** 需求提出人 */
    private String demandProposer;

    /** 需求提出时间 */
    private Date demandProposeTime;

    /** 需求状态 */
    private String demandStatus;

    /** 需求开发完成时间 */
    private Date demandFinishTime;

    /** 需求开发完成说明 */
    private String demandFinishRemark;

    /** TB单号 */
    private String tbBillno;

    /** 记录创建时间 */
    private Date createTime;

    /** 记录更新时间 */
    private Date updateTime;

    private static final long serialVersionUID = 1L;
}
