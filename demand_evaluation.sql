-- ============================================================
-- 需求评估导入表 demand_evaluation - 完整SQL脚本
-- 数据库：tp_music
-- 生成时间：2026-07-09
-- ============================================================

-- 1. 删除已有表
DROP TABLE IF EXISTS `demand_evaluation`;

-- 2. 创建表
CREATE TABLE `demand_evaluation` (
  `id`                     BIGINT        NOT NULL AUTO_INCREMENT COMMENT '主键ID',
  `seq_no`                 INT           NOT NULL                COMMENT '序号',
  `program_code`           VARCHAR(50)   NOT NULL                COMMENT '程序代号',
  `program_name`           VARCHAR(200)  NOT NULL                COMMENT '程序名称',
  `billing_hours_customer` DECIMAL(10,2) DEFAULT NULL            COMMENT '计费时数-客户',
  `dispatch_hours`         DECIMAL(10,2) DEFAULT NULL            COMMENT '软代派工时数',
  `system_code`            VARCHAR(50)   NOT NULL                COMMENT '系统代号',
  `spec_desc`              TEXT          NOT NULL                COMMENT '规格说明',
  `custom_type`            VARCHAR(50)   NOT NULL                COMMENT '客制类型',
  `difficulty_level`       VARCHAR(50)   NOT NULL                COMMENT '难度等级',
  `dev_category`           VARCHAR(100)  NOT NULL                COMMENT '开发分类',
  `integrate_product_code` VARCHAR(100)  DEFAULT NULL            COMMENT '集成产品编号',
  `urgent_flag`            VARCHAR(10)   DEFAULT NULL            COMMENT '急单',
  `remark`                 VARCHAR(500)  DEFAULT NULL            COMMENT '备注',
  `demand_proposer`        VARCHAR(100)  DEFAULT NULL            COMMENT '需求提出人',
  `demand_propose_time`    DATETIME      DEFAULT NULL            COMMENT '需求提出时间',
  `demand_status`          VARCHAR(50)   DEFAULT NULL            COMMENT '需求状态',
  `demand_finish_time`     DATETIME      DEFAULT NULL            COMMENT '需求开发完成时间',
  `demand_finish_remark`   VARCHAR(500)  DEFAULT NULL            COMMENT '需求开发完成说明',
  `create_time`            DATETIME      DEFAULT CURRENT_TIMESTAMP                    COMMENT '记录创建时间',
  `update_time`            DATETIME      DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '记录更新时间',
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='需求评估导入表';

-- 3. 插入数据（共12条）
INSERT INTO `demand_evaluation` (
  `seq_no`, `program_code`, `program_name`,
  `billing_hours_customer`, `dispatch_hours`,
  `system_code`, `spec_desc`,
  `custom_type`, `difficulty_level`, `dev_category`,
  `integrate_product_code`, `urgent_flag`, `remark`,
  `demand_proposer`, `demand_propose_time`, `demand_status`,
  `demand_finish_time`, `demand_finish_remark`
) VALUES
-- 第1条：物料主档维护作业
(1, 'aimm200', '物料主档维护作业', NULL, NULL,
 'AIM',
 '需求逻辑：\n1、aimm200新增字段，字段名称"报价物料"\n2、PLM调标准接口抛ERP\n（adzi180、s_aws_plm_ima调整）',
 '2.修改标准', '2:二级(初级)', '702:服务端_生单接口',
 'G36: SRM_正远数智', 'N', NULL,
 NULL, NULL, NULL, NULL, NULL),

-- 第2条：集成服务端检测工具
(2, 'awsq990', '集成服务端检测工具', NULL, NULL,
 'AWS',
 '需求逻辑：\n1、awsq990结束时间下方增加"关键字"字段\n2、增加查询时按照关键字搜索（主要判断空）',
 '2.修改标准', '2:二级(初级)', '201:建档(通用)',
 NULL, 'N', NULL,
 NULL, NULL, NULL, NULL, NULL),

-- 第3条：物料主档-视图
(3, 'aimm200', '物料主档-视图', NULL, NULL,
 'AIM',
 '物料主档视图：all_aimm200uc_v\n视图字段：\nimaacrtid 创建人账号\nimaacrtdp 资料录入部门\nimaacrtdt 资料创建日\nimaamoddt 最近更改日\nimaa001 料号\nimaal003 品名\nimaal004 规格\nimaa009 产品分类\nimaa004 料件类别\nimaf091 默认库位\nimaaent 企业编号\nimaa006 基础单位\nimaf015 参考单位\nimad003 基础单位数量\nimad005 参考单位数量\nimaf166 采购超交率\nimaa010 生命周期\nimaf142 采购员说明档\nimaf026 安全库存',
 '1.新增', '2:二级(初级)', '201:建档(通用)',
 NULL, 'N', NULL,
 NULL, NULL, NULL, NULL, NULL),

-- 第4条：产品分类-视图
(4, 'aimi010', '产品分类-视图', NULL, NULL,
 'AIM',
 '产品分类视图：all_aimi010uc_v\nrtax001 分类编码\nrtaxl003 分类名称\nrtax003 父级分类编码值为空表示一级分类\nrtaxstus 启用否',
 '1.新增', '2:二级(初级)', '201:建档(通用)',
 NULL, 'N', NULL,
 NULL, NULL, NULL, NULL, NULL),

-- 第5条：请购单-视图
(5, 'apmt400', '请购单-视图', NULL, NULL,
 'APM',
 '请购单视图：all_apmt400uc_v\npmdacrtid 资料录入者\npmdacrtdp 资料录入部门\npmdacrtid 创建人\npmdacrtdp 创建部门\npmdb001 请购单号\npmdb002 请购项次\npmdb003 请购日\npmaa001 料号\npmaal003 品名\npmaal004 规格\npmdc001 请购数量\npmdc002 请购单位\npmdb004 预计到货日\npmdc003 预估单价\npmdc004 预估总价\npmdb005 请购状态\npmdc005 币别\npmdc006 汇率\npmdc007 请购类型\npmdb006 来源单号\npmdb007 来源项次\npmdc008 需求者\npmdc009 需求部门\npmdc010 项目编码\npmdc011 备注',
 '1.新增', '2:二级(初级)', '201:建档(通用)',
 NULL, 'N', NULL,
 NULL, NULL, NULL, NULL, NULL),

-- 第6条：请购单变更单-视图
(6, 'apmt410', '请购单变更单-视图', NULL, NULL,
 'APM',
 '请购变更单视图：all_apmt410uc_v\npmfacrtid 资料录入者\npmfacrtdp 资料录入部门\npmfacrtid 创建人\npmfacrtdp 创建部门\npmfb001 请购变更单号\npmfb002 请购变更项次\npmfb003 请购变更日\npmaa001 料号\npmaal003 品名\npmaal004 规格\npmfc001 变更数量\npmfc002 变更单位\npmfb004 预计到货日\npmfc003 预估单价\npmfc004 预估总价\npmfb005 变更状态\npmfc005 币别\npmfc006 汇率\npmfc007 变更类型\npmfb006 来源单号\npmfb007 来源项次\npmfc008 需求者\npmfc009 需求部门\npmfc010 项目编码\npmfc011 备注',
 '1.新增', '2:二级(初级)', '201:建档(通用)',
 NULL, 'N', NULL,
 NULL, NULL, NULL, NULL, NULL),

-- 第7条：请购单核准作业
(7, 'apmt420', '请购单核准作业', NULL, NULL,
 'APM',
 '请购单核准：\n1、请购单核准流程\n2、核准后更新请购单状态',
 '1.新增', '2:二级(初级)', '201:建档(通用)',
 NULL, 'N', NULL,
 NULL, NULL, NULL, NULL, NULL),

-- 第8条：采购单作业
(8, 'apmt500', '采购单作业', NULL, NULL,
 'APM',
 '接口方向：SRM->T100\n接口字段：接收采购单数据接口\n1、采购单主文件\n2、采购单明细',
 '1.新增', '2:二级(初级)', '702:服务端_生单接口',
 'G36: SRM_正远数智', 'N', NULL,
 NULL, NULL, NULL, NULL, NULL),

-- 第9条：采购单-视图
(9, 'apmt500', '采购单-视图', NULL, NULL,
 'APM',
 '采购单视图：all_apmt500uc_v\npmacrtid 资料录入者\npmacrtdp 资料录入部门\npmaa001 供应商编码\npmaal003 供应商名称\npmab001 采购单号\npmab002 采购项次\npmab003 采购日\npmaa001 料号\npmaal003 品名\npmaal004 规格\npmac001 采购数量\npmac002 采购单位\npmab004 预计到货日\npmac003 单价\npmac004 总价\npmab005 采购状态\npmac005 币别\npmac006 汇率\npmac007 采购类型\npmab006 来源单号\npmab007 来源项次\npmac008 需求者\npmac009 需求部门\npmac010 项目编码\npmac011 备注',
 '1.新增', '2:二级(初级)', '702:服务端_生单接口',
 'G36: SRM_正远数智', 'N', NULL,
 NULL, NULL, NULL, NULL, NULL),

-- 第10条：采购变更单作业
(10, 'apmt510', '采购变更单作业', NULL, NULL,
 'APM',
 '接口方向：SRM->T100\n接口字段：接收采购变更单数据接口，且自动更新apmt500对应的版次',
 '1.新增', '2:二级(初级)', '702:服务端_生单接口',
 'G36: SRM_正远数智', 'N', NULL,
 NULL, NULL, NULL, NULL, NULL),

-- 第11条：采购收货单作业
(11, 'apmt520', '采购收货单作业', NULL, NULL,
 'APM',
 '接口方向：SRM->T100\n接口字段：接收采购收货数据，记录SRM源送货单号',
 '1.新增', '2:二级(初级)', '702:服务端_生单接口',
 'G36: SRM_正远数智', 'N', NULL,
 NULL, NULL, NULL, NULL, NULL),

-- 第12条：供应商发票作业
(12, 'aapt110', '供应商发票作业', NULL, NULL,
 'AAP',
 '接口方向：SRM->T100\n接口字段：接收采购发票数据',
 '1.新增', '2:二级(初级)', '702:服务端_生单接口',
 'G36: SRM_正远数智', 'N', NULL,
 NULL, NULL, NULL, NULL, NULL);
