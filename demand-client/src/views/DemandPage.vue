<template>
  <div class="container">
    <div class="handle-box">
      <el-button @click="deleteAll">批量删除</el-button>
      <el-input v-model="searchWord" placeholder="筛选需求" style="width: 200px; margin-right: 10px;"></el-input>
      <el-button type="primary" @click="addDialogVisible = true">添加需求</el-button>
      <el-upload
        action=""
        :auto-upload="false"
        :show-file-list="false"
        accept=".xlsx,.xls"
        :on-change="handleImport"
      >
        <el-button type="success">需求清单导入</el-button>
      </el-upload>
      <el-button type="warning" @click="exportFieldMapping">需求清单字段对照表导出</el-button>
      <el-button type="info" @click="exportAiEnhanced" :loading="aiLoading">AI 智能分析导出</el-button>
      <el-button type="primary" @click="exportApiMatch" :loading="apiMatchLoading">API对照匹配导出</el-button>
    </div>

    <el-table height="550px" border size="small" :data="data" @selection-change="handleSelectionChange">
      <el-table-column type="selection" width="40" align="center"></el-table-column>
      <el-table-column label="ID" prop="id" width="50" align="center"></el-table-column>
      <el-table-column label="序号" prop="seqNo" width="60" align="center"></el-table-column>
      <el-table-column label="程序代号" prop="programCode" width="100" align="center"></el-table-column>
      <el-table-column label="程序名称" prop="programName" width="150"></el-table-column>
      <el-table-column label="系统" prop="systemCode" width="60" align="center"></el-table-column>
      <el-table-column label="规格说明" prop="specDesc" min-width="200">
        <template v-slot="scope">
          <el-popover placement="left" trigger="hover" width="400">
            <template #reference>
              <div class="text-ellipsis">{{ scope.row.specDesc }}</div>
            </template>
            <div style="white-space: pre-wrap; max-height: 300px; overflow-y: auto;">{{ scope.row.specDesc }}</div>
          </el-popover>
        </template>
      </el-table-column>
      <el-table-column label="客制类型" prop="customType" width="90" align="center"></el-table-column>
      <el-table-column label="难度" prop="difficultyLevel" width="100" align="center"></el-table-column>
      <el-table-column label="开发分类" prop="devCategory" width="140" align="center"></el-table-column>
      <el-table-column label="集成产品" prop="integrateProductCode" width="130" align="center"></el-table-column>
      <el-table-column label="急单" prop="urgentFlag" width="50" align="center"></el-table-column>
      <el-table-column label="提出人" prop="demandProposer" width="80" align="center"></el-table-column>
      <el-table-column label="状态" prop="demandStatus" width="80" align="center"></el-table-column>
      <el-table-column label="TB单号" prop="tbBillno" width="120" align="center"></el-table-column>
      <el-table-column label="操作" width="140" align="center" fixed="right">
        <template v-slot="scope">
          <el-button @click="editRow(scope.row)">编辑</el-button>
          <el-button type="danger" @click="deleteRow(scope.row.id)">删除</el-button>
        </template>
      </el-table-column>
    </el-table>
    <el-pagination
      class="pagination"
      background
      layout="total, prev, pager, next"
      :current-page="currentPage"
      :page-size="pageSize"
      :total="tableData.length"
      @current-change="handleCurrentChange"
    ></el-pagination>
  </div>

  <!-- 添加弹窗 -->
  <el-dialog title="添加需求" v-model="addDialogVisible" width="700px">
    <el-form :model="addForm" label-width="120px">
      <el-row :gutter="20">
        <el-col :span="12"><el-form-item label="序号"><el-input-number v-model="addForm.seqNo" :min="1" style="width:100%"></el-input-number></el-form-item></el-col>
        <el-col :span="12"><el-form-item label="系统代号"><el-input v-model="addForm.systemCode"></el-input></el-form-item></el-col>
      </el-row>
      <el-row :gutter="20">
        <el-col :span="12"><el-form-item label="程序代号"><el-input v-model="addForm.programCode"></el-input></el-form-item></el-col>
        <el-col :span="12"><el-form-item label="程序名称"><el-input v-model="addForm.programName"></el-input></el-form-item></el-col>
      </el-row>
      <el-form-item label="规格说明"><el-input type="textarea" v-model="addForm.specDesc" :rows="4"></el-input></el-form-item>
      <el-row :gutter="20">
        <el-col :span="12">
          <el-form-item label="客制类型">
            <el-select v-model="addForm.customType" style="width:100%">
              <el-option label="1.新增" value="1.新增"></el-option>
              <el-option label="2.修改标准" value="2.修改标准"></el-option>
              <el-option label="3.小修客制" value="3.小修客制"></el-option>
              <el-option label="4.版更" value="4.版更"></el-option>
            </el-select>
          </el-form-item>
        </el-col>
        <el-col :span="12">
          <el-form-item label="难度等级">
            <el-select v-model="addForm.difficultyLevel" style="width:100%">
              <el-option label="1:一级(入门)" value="1:一级(入门)"></el-option>
              <el-option label="2:二级(初级)" value="2:二级(初级)"></el-option>
              <el-option label="3:三级(中级)" value="3:三级(中级)"></el-option>
              <el-option label="4:四级(高级)" value="4:四级(高级)"></el-option>
              <el-option label="5:五级(资级)" value="5:五级(资级)"></el-option>
            </el-select>
          </el-form-item>
        </el-col>
      </el-row>
      <el-row :gutter="20">
        <el-col :span="12"><el-form-item label="开发分类"><el-input v-model="addForm.devCategory"></el-input></el-form-item></el-col>
        <el-col :span="12"><el-form-item label="集成产品编号"><el-input v-model="addForm.integrateProductCode"></el-input></el-form-item></el-col>
      </el-row>
      <el-row :gutter="20">
        <el-col :span="12"><el-form-item label="计费时数"><el-input-number v-model="addForm.billingHoursCustomer" :precision="2" :min="0" style="width:100%"></el-input-number></el-form-item></el-col>
        <el-col :span="12"><el-form-item label="派工时数"><el-input-number v-model="addForm.dispatchHours" :precision="2" :min="0" style="width:100%"></el-input-number></el-form-item></el-col>
      </el-row>
      <el-row :gutter="20">
        <el-col :span="12">
          <el-form-item label="急单">
            <el-select v-model="addForm.urgentFlag" style="width:100%">
              <el-option label="Y" value="Y"></el-option><el-option label="N" value="N"></el-option>
            </el-select>
          </el-form-item>
        </el-col>
        <el-col :span="12"><el-form-item label="备注"><el-input v-model="addForm.remark"></el-input></el-form-item></el-col>
      </el-row>
      <el-row :gutter="20">
        <el-col :span="12"><el-form-item label="需求提出人"><el-input v-model="addForm.demandProposer"></el-input></el-form-item></el-col>
        <el-col :span="12"><el-form-item label="需求状态"><el-input v-model="addForm.demandStatus"></el-input></el-form-item></el-col>
      </el-row>
      <el-row :gutter="20">
        <el-col :span="12"><el-form-item label="TB单号"><el-input v-model="addForm.tbBillno"></el-input></el-form-item></el-col>
      </el-row>
      <el-form-item label="完成说明"><el-input type="textarea" v-model="addForm.demandFinishRemark" :rows="2"></el-input></el-form-item>
    </el-form>
    <template #footer>
      <el-button @click="addDialogVisible = false">取消</el-button>
      <el-button type="primary" @click="addDemand">确定</el-button>
    </template>
  </el-dialog>

  <!-- 编辑弹窗 -->
  <el-dialog title="编辑需求" v-model="editDialogVisible" width="700px">
    <el-form :model="editForm" label-width="120px">
      <el-row :gutter="20">
        <el-col :span="12"><el-form-item label="序号"><el-input-number v-model="editForm.seqNo" :min="1" style="width:100%"></el-input-number></el-form-item></el-col>
        <el-col :span="12"><el-form-item label="系统代号"><el-input v-model="editForm.systemCode"></el-input></el-form-item></el-col>
      </el-row>
      <el-row :gutter="20">
        <el-col :span="12"><el-form-item label="程序代号"><el-input v-model="editForm.programCode"></el-input></el-form-item></el-col>
        <el-col :span="12"><el-form-item label="程序名称"><el-input v-model="editForm.programName"></el-input></el-form-item></el-col>
      </el-row>
      <el-form-item label="规格说明"><el-input type="textarea" v-model="editForm.specDesc" :rows="4"></el-input></el-form-item>
      <el-row :gutter="20">
        <el-col :span="12">
          <el-form-item label="客制类型">
            <el-select v-model="editForm.customType" style="width:100%">
              <el-option label="1.新增" value="1.新增"></el-option><el-option label="2.修改标准" value="2.修改标准"></el-option>
              <el-option label="3.小修客制" value="3.小修客制"></el-option><el-option label="4.版更" value="4.版更"></el-option>
            </el-select>
          </el-form-item>
        </el-col>
        <el-col :span="12">
          <el-form-item label="难度等级">
            <el-select v-model="editForm.difficultyLevel" style="width:100%">
              <el-option label="1:一级(入门)" value="1:一级(入门)"></el-option><el-option label="2:二级(初级)" value="2:二级(初级)"></el-option>
              <el-option label="3:三级(中级)" value="3:三级(中级)"></el-option><el-option label="4:四级(高级)" value="4:四级(高级)"></el-option>
              <el-option label="5:五级(资级)" value="5:五级(资级)"></el-option>
            </el-select>
          </el-form-item>
        </el-col>
      </el-row>
      <el-row :gutter="20">
        <el-col :span="12"><el-form-item label="开发分类"><el-input v-model="editForm.devCategory"></el-input></el-form-item></el-col>
        <el-col :span="12"><el-form-item label="集成产品编号"><el-input v-model="editForm.integrateProductCode"></el-input></el-form-item></el-col>
      </el-row>
      <el-row :gutter="20">
        <el-col :span="12"><el-form-item label="计费时数"><el-input-number v-model="editForm.billingHoursCustomer" :precision="2" :min="0" style="width:100%"></el-input-number></el-form-item></el-col>
        <el-col :span="12"><el-form-item label="派工时数"><el-input-number v-model="editForm.dispatchHours" :precision="2" :min="0" style="width:100%"></el-input-number></el-form-item></el-col>
      </el-row>
      <el-row :gutter="20">
        <el-col :span="12">
          <el-form-item label="急单">
            <el-select v-model="editForm.urgentFlag" style="width:100%">
              <el-option label="Y" value="Y"></el-option><el-option label="N" value="N"></el-option>
            </el-select>
          </el-form-item>
        </el-col>
        <el-col :span="12"><el-form-item label="备注"><el-input v-model="editForm.remark"></el-input></el-form-item></el-col>
      </el-row>
      <el-row :gutter="20">
        <el-col :span="12"><el-form-item label="需求提出人"><el-input v-model="editForm.demandProposer"></el-input></el-form-item></el-col>
        <el-col :span="12"><el-form-item label="需求状态"><el-input v-model="editForm.demandStatus"></el-input></el-form-item></el-col>
      </el-row>
      <el-row :gutter="20">
        <el-col :span="12"><el-form-item label="TB单号"><el-input v-model="editForm.tbBillno"></el-input></el-form-item></el-col>
        <el-col :span="12"><el-form-item label="提出时间"><el-date-picker v-model="editForm.demandProposeTime" type="datetime" style="width:100%"></el-date-picker></el-form-item></el-col>
        <el-col :span="12"><el-form-item label="完成时间"><el-date-picker v-model="editForm.demandFinishTime" type="datetime" style="width:100%"></el-date-picker></el-form-item></el-col>
      </el-row>
      <el-form-item label="完成说明"><el-input type="textarea" v-model="editForm.demandFinishRemark" :rows="2"></el-input></el-form-item>
    </el-form>
    <template #footer>
      <el-button @click="editDialogVisible = false">取消</el-button>
      <el-button type="primary" @click="saveEdit">确定</el-button>
    </template>
  </el-dialog>

  <!-- 删除确认 -->
  <yin-del-dialog :delVisible="delVisible" @confirm="confirm" @cancelRow="delVisible = $event"></yin-del-dialog>
</template>

<script lang="ts">
import { defineComponent, watch, ref, reactive, computed, onMounted } from "vue";
import { ElMessage } from "element-plus";
import { HttpManager } from "@/api";
import YinDelDialog from "@/components/dialog/YinDelDialog.vue";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

export default defineComponent({
  name: "DemandPage",
  components: { YinDelDialog },
  setup() {
    const tableData = ref([]);
    const tempDate = ref([]);
    const pageSize = ref(10);
    const currentPage = ref(1);
    const aiLoading = ref(false);
    const apiMatchLoading = ref(false);

    const data = computed(() =>
      tableData.value.slice(
        (currentPage.value - 1) * pageSize.value,
        currentPage.value * pageSize.value
      )
    );

    // 搜索
    const searchWord = ref("");
    watch(searchWord, () => {
      if (!searchWord.value) {
        tableData.value = tempDate.value;
      } else {
        tableData.value = tempDate.value.filter(
          (item: any) =>
            (item.programName || "").includes(searchWord.value) ||
            (item.programCode || "").includes(searchWord.value) ||
            (item.specDesc || "").includes(searchWord.value) ||
            (item.tbBillno || "").includes(searchWord.value)
        );
      }
      currentPage.value = 1;
    });

    // 加载数据
    async function getData() {
      console.log("开始加载数据...");
      try {
        const result = (await HttpManager.getAllDemand()) as any;
        console.log("API返回结果:", result);
        if (result && result.data) {
          tableData.value = result.data;
          tempDate.value = result.data;
          console.log("数据加载成功, 共", result.data.length, "条");
        } else {
          console.warn("API返回数据为空:", result);
        }
        currentPage.value = 1;
      } catch (e) {
        console.error("获取数据失败:", e);
        ElMessage.error("获取数据失败");
      }
    }

    onMounted(() => {
      getData();
    });

    function handleCurrentChange(val: number) {
      currentPage.value = val;
    }

    // =================== 添加 ===================
    const addDialogVisible = ref(false);
    const defaultForm = {
      seqNo: 1,
      programCode: "",
      programName: "",
      billingHoursCustomer: null as number | null,
      dispatchHours: null as number | null,
      systemCode: "",
      specDesc: "",
      customType: "1.新增",
      difficultyLevel: "2:二级(初级)",
      devCategory: "",
      integrateProductCode: "",
      urgentFlag: "N",
      remark: "",
      demandProposer: "",
      demandStatus: "",
      demandFinishRemark: "",
      tbBillno: ""
    };
    const addForm = reactive({ ...defaultForm });

    async function addDemand() {
      try {
        const result = (await HttpManager.addDemand(addForm)) as any;
        if (result && result.success) {
          ElMessage.success(result.message || "添加成功");
          getData();
          addDialogVisible.value = false;
          Object.assign(addForm, { ...defaultForm });
        } else {
          ElMessage.error(result?.message || "添加失败");
        }
      } catch (e) {
        ElMessage.error("添加失败");
      }
    }

    // =================== 编辑 ===================
    const editDialogVisible = ref(false);
    const editForm = reactive({
      id: null as number | null,
      ...defaultForm,
      demandProposeTime: null,
      demandFinishTime: null
    });

    function editRow(row: any) {
      Object.assign(editForm, {
        id: row.id,
        seqNo: row.seqNo,
        programCode: row.programCode || "",
        programName: row.programName || "",
        billingHoursCustomer: row.billingHoursCustomer,
        dispatchHours: row.dispatchHours,
        systemCode: row.systemCode || "",
        specDesc: row.specDesc || "",
        customType: row.customType || "",
        difficultyLevel: row.difficultyLevel || "",
        devCategory: row.devCategory || "",
        integrateProductCode: row.integrateProductCode || "",
        urgentFlag: row.urgentFlag || "N",
        remark: row.remark || "",
        demandProposer: row.demandProposer || "",
        demandStatus: row.demandStatus || "",
        demandFinishRemark: row.demandFinishRemark || "",
        demandProposeTime: row.demandProposeTime,
        demandFinishTime: row.demandFinishTime,
        tbBillno: row.tbBillno || ""
      });
      editDialogVisible.value = true;
    }

    async function saveEdit() {
      try {
        const result = (await HttpManager.updateDemand(editForm)) as any;
        if (result && result.success) {
          ElMessage.success(result.message || "更新成功");
          getData();
          editDialogVisible.value = false;
        } else {
          ElMessage.error(result?.message || "更新失败");
        }
      } catch (e) {
        ElMessage.error("更新失败");
      }
    }

    // =================== 删除 ===================
    const idx = ref(-1);
    const multipleSelection = ref([]);
    const delVisible = ref(false);

    async function confirm() {
      try {
        const result = (await HttpManager.deleteDemand(idx.value)) as any;
        if (result && result.success) {
          ElMessage.success(result.message || "删除成功");
          getData();
        } else {
          ElMessage.error(result?.message || "删除失败");
        }
      } catch (e) {
        ElMessage.error("删除失败");
      }
      delVisible.value = false;
    }

    function deleteRow(id: number) {
      idx.value = id;
      delVisible.value = true;
    }

    function handleSelectionChange(val: any[]) {
      multipleSelection.value = val;
    }

    function deleteAll() {
      for (const item of multipleSelection.value) {
        deleteRow(item.id);
        confirm();
      }
      multipleSelection.value = [];
    }

    // =================== Excel 导入 ===================
    async function handleImport(file: any) {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: "array" });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(firstSheet);

          if (!jsonData || jsonData.length === 0) {
            ElMessage.warning("Excel 文件中没有数据");
            return;
          }

          // 字段映射（Excel列名 -> 数据库字段）
          // 支持带 (*) 后缀的列名，以及多种列名别名
          const fieldMapping: Record<string, string> = {
            "序号": "seqNo",
            "序号(*)": "seqNo",
            "程序代号": "programCode",
            "程序代号(*)": "programCode",
            "程序名称": "programName",
            "程序名称(*)": "programName",
            "系统代号": "systemCode",
            "系统代号(*)": "systemCode",
            "规格说明": "specDesc",
            "规格说明(*)": "specDesc",
            "客制类型": "customType",
            "客制类型(*)": "customType",
            "难度等级": "difficultyLevel",
            "难度等级(*)": "difficultyLevel",
            "开发分类": "devCategory",
            "开发分类(*)": "devCategory",
            "集成产品编号": "integrateProductCode",
            "计费时数-客户": "billingHoursCustomer",
            "计费时数-客户(*)": "billingHoursCustomer",
            "计费时数": "billingHoursCustomer",
            "软代派工时数": "dispatchHours",
            "软代派工时数(*)": "dispatchHours",
            "派工时数": "dispatchHours",
            "急单": "urgentFlag",
            "备注": "remark",
            "TB单号": "tbBillno",
            "需求提出人": "demandProposer",
            "需求状态": "demandStatus",
            "完成说明": "demandFinishRemark"
          };

          let successCount = 0;
          let failCount = 0;

          for (const row of jsonData) {
            const mappedRow: Record<string, any> = {};
            for (const [cnKey, enKey] of Object.entries(fieldMapping)) {
              if (row[cnKey] !== undefined) {
                mappedRow[enKey] = row[cnKey];
              }
            }
            // 设置默认值
            if (!mappedRow.seqNo) mappedRow.seqNo = 1;
            if (!mappedRow.customType) mappedRow.customType = "1.新增";
            if (!mappedRow.difficultyLevel) mappedRow.difficultyLevel = "2:二级(初级)";
            if (!mappedRow.urgentFlag) mappedRow.urgentFlag = "N";

            try {
              const result = (await HttpManager.addDemand(mappedRow)) as any;
              if (result && result.success) {
                successCount++;
              } else {
                failCount++;
              }
            } catch {
              failCount++;
            }
          }

          ElMessage.success(`导入完成：成功 ${successCount} 条，失败 ${failCount} 条`);
          getData();
        } catch (err) {
          console.error("导入失败:", err);
          ElMessage.error("导入失败，请检查文件格式");
        }
      };
      reader.readAsArrayBuffer(file.raw);
    }

    // =================== 字段对照表导出 ===================
    function exportFieldMapping() {
      // 字段对照表数据
      const mappingData = [
        { "Excel列名": "序号", "数据库字段": "seqNo", "类型": "数字", "说明": "需求序号，从1开始递增", "示例": "1" },
        { "Excel列名": "程序代号", "数据库字段": "programCode", "类型": "文本", "说明": "程序/模块代号", "示例": "aimm200" },
        { "Excel列名": "程序名称", "数据库字段": "programName", "类型": "文本", "说明": "程序/模块名称", "示例": "物料主档维护作业" },
        { "Excel列名": "系统代号", "数据库字段": "systemCode", "类型": "文本", "说明": "所属系统代号", "示例": "AIM" },
        { "Excel列名": "规格说明", "数据库字段": "specDesc", "类型": "文本", "说明": "需求详细规格描述", "示例": "新增字段xxx" },
        { "Excel列名": "客制类型", "数据库字段": "customType", "类型": "枚举", "说明": "1.新增/2.修改标准/3.小修客制/4.版更", "示例": "1.新增" },
        { "Excel列名": "难度等级", "数据库字段": "difficultyLevel", "类型": "枚举", "说明": "1:一级(入门)/2:二级(初级)/3:三级(中级)/4:四级(高级)/5:五级(资级)", "示例": "2:二级(初级)" },
        { "Excel列名": "开发分类", "数据库字段": "devCategory", "类型": "文本", "说明": "开发分类编码", "示例": "201:建档(通用)" },
        { "Excel列名": "集成产品编号", "数据库字段": "integrateProductCode", "类型": "文本", "说明": "集成产品编号", "示例": "G36: SRM_正远数智" },
        { "Excel列名": "计费时数", "数据库字段": "billingHoursCustomer", "类型": "数字", "说明": "客户计费时数", "示例": "8.5" },
        { "Excel列名": "派工时数", "数据库字段": "dispatchHours", "类型": "数字", "说明": "派工时数", "示例": "6.0" },
        { "Excel列名": "急单", "数据库字段": "urgentFlag", "类型": "枚举", "说明": "Y=是/N=否", "示例": "N" },
        { "Excel列名": "备注", "数据库字段": "remark", "类型": "文本", "说明": "备注信息", "示例": "注意事项" },
        { "Excel列名": "需求提出人", "数据库字段": "demandProposer", "类型": "文本", "说明": "需求提出人姓名", "示例": "张三" },
        { "Excel列名": "需求状态", "数据库字段": "demandStatus", "类型": "文本", "说明": "需求当前状态", "示例": "开发中" },
        { "Excel列名": "完成说明", "数据库字段": "demandFinishRemark", "类型": "文本", "说明": "开发完成说明", "示例": "已完成" },
        { "Excel列名": "TB单号", "数据库字段": "tbBillno", "类型": "文本", "说明": "TB单号", "示例": "TB202607001" }
      ];

      // 创建工作簿
      const ws = XLSX.utils.json_to_sheet(mappingData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "字段对照表");

      // 设置列宽
      ws["!cols"] = [
        { wch: 15 }, // Excel列名
        { wch: 22 }, // 数据库字段
        { wch: 8 },  // 类型
        { wch: 50 }, // 说明
        { wch: 20 }  // 示例
      ];

      // 导出
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbout], { type: "application/octet-stream" });
      saveAs(blob, "需求清单字段对照表.xlsx");
    }

    // =================== AI 智能分析导出 ===================
    // 先获取 AI 分析结果，再调用完整导出（产出包含字段对照表 + AI 分析的综合 Excel）
    async function exportAiEnhanced() {
      const exportData: any[] = multipleSelection.value.length > 0 
        ? multipleSelection.value 
        : tableData.value;
      
      if (exportData.length === 0) {
        ElMessage.warning("没有可导出的数据");
        return;
      }
      
      aiLoading.value = true;
      try {
        // 1. 获取 AI 分析结果
        const aiResult = (await HttpManager.aiExportEnhanced(exportData)) as any;
        
        if (!aiResult || !aiResult.success || !aiResult.data) {
          ElMessage.error("AI 分析失败: " + (aiResult?.message || "未知错误"));
          return;
        }

        const aiData = aiResult.data;
        
        // 2. 从 AI 匹配结果构建 demandAnalyses
        const demandAnalyses: any[] = [];
        const matches = aiData.matches || [];
        for (const match of matches) {
          demandAnalyses.push({
            id: match.id,
            programCode: exportData.find((d: any) => d.id === match.id)?.programCode || "",
            programName: exportData.find((d: any) => d.id === match.id)?.programName || "",
            classification: {
              customType: match.customType || "",
              difficultyLevel: ""
            },
            hourEstimate: {
              billingHoursCustomer: "",
              dispatchHours: "",
              reason: match.reason || ""
            },
            optimization: {
              completenessScore: "",
              suggestions: []
            }
          });
        }
        
        // 3. 调用完整导出接口（产出包含字段对照表 + AI 分析的综合 Excel）
        const exportPayload = {
          demands: exportData,
          aiData: {
            summary: aiData.summary || {},
            demandAnalyses: demandAnalyses
          }
        };
        
        const response = (await HttpManager.exportWithApiMatch(exportPayload)) as any;
        
        if (!response || !response.data) {
          ElMessage.error("导出失败: 服务器返回空数据");
          return;
        }
        
        // 4. 下载文件
        const blob = response.data;
        if (blob.type && blob.type.includes('application/json')) {
          const text = await blob.text();
          ElMessage.error("导出失败: " + text);
          return;
        }
        
        const excelBlob = new Blob([blob], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        
        const url = window.URL.createObjectURL(excelBlob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `需求清单_AI智能分析_${new Date().toISOString().slice(0, 10)}.xlsx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
        
        ElMessage.success("AI 智能分析导出成功（含字段对照表）");
      } catch (err) {
        console.error("AI 导出失败:", err);
        ElMessage.error("AI 导出失败: " + (err as Error).message);
      } finally {
        aiLoading.value = false;
      }
    }

    // =================== API 对照匹配导出 ===================
    async function exportApiMatch() {
      // 获取导出数据：优先使用选中行，否则使用全部数据
      const exportData: any[] = multipleSelection.value.length > 0 
        ? multipleSelection.value 
        : tableData.value;
      
      if (exportData.length === 0) {
        ElMessage.warning("没有可导出的数据");
        return;
      }
      
      apiMatchLoading.value = true;
      try {
        const response = (await HttpManager.exportWithApiMatch(exportData)) as any;
        
        if (!response || !response.data) {
          ElMessage.error("导出失败: 服务器返回空数据");
          return;
        }
        
        // 检查响应是否为错误信息（blob 可能是 JSON 错误）
        const blob = response.data;
        if (blob.type && blob.type.includes('application/json')) {
          const text = await blob.text();
          ElMessage.error("导出失败: " + text);
          return;
        }
        
        // 创建 Blob 并下载
        const excelBlob = new Blob([blob], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        
        const url = window.URL.createObjectURL(excelBlob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `需求清单_API匹配结果_${new Date().toISOString().slice(0, 10)}.xlsx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
        
        ElMessage.success("API 对照匹配导出成功");
      } catch (err) {
        console.error("导出失败:", err);
        ElMessage.error("导出失败: " + (err as Error).message);
      } finally {
        apiMatchLoading.value = false;
      }
    }

    return {
      searchWord,
      data,
      tableData,
      addDialogVisible,
      addForm,
      editDialogVisible,
      editForm,
      delVisible,
      pageSize,
      currentPage,
      handleCurrentChange,
      addDemand,
      editRow,
      saveEdit,
      deleteRow,
      deleteAll,
      confirm,
      handleSelectionChange,
      handleImport,
      exportFieldMapping,
      aiLoading,
      exportAiEnhanced,
      apiMatchLoading,
      exportApiMatch
    };
  }
});
</script>

<style scoped>
.text-ellipsis {
  overflow: hidden;
  text-overflow: ellipsis;
  display: -webkit-box;
  -webkit-line-clamp: 3;
  -webkit-box-orient: vertical;
  max-height: 60px;
  word-break: break-all;
}
</style>
