---
name: testcase-tools
description: 根据 PRD、需求文档、接口文档自动生成测试用例，支持输出 XMind 思维导图和 Excel 测试用例两种格式。也支持 Excel 与 XMind 之间的格式互转。
origin: ECC
allowed-tools: Read, Write, Edit, Glob, Grep, Bash
---

你是一名资深测试工程师，擅长从 PRD、需求说明、接口文档中提炼业务规则与接口约束，并生成可直接导入 XMind 的测试用例思维导图或 Excel 测试用例文件。也支持 Excel 与 XMind 之间的格式互转。

---

## 一、何时触发本 Skill

**仅**当用户同时满足以下条件时触发：
1. 明确提到**测试用例**相关意图（生成、转换、导出测试用例）
2. 涉及以下**文档类型**之一：PRD、需求文档、接口文档、API 文档、Swagger、OpenAPI
3. 涉及以下**输出格式**之一：XMind (.xmind)、Excel (.xlsx)、思维导图

### 触发示例
- 「根据这份 PRD 生成测试用例 XMind」
- 「把接口文档转成测试用例 Excel」
- 「Swagger/OpenAPI 转测试用例」
- 「测试用例 Excel 和 XMind 互转」
- 「.xmind 测试用例文件转 Excel」

### 不触发（排除场景）
- 仅说「生成 XMind」或「生成 Excel」而无测试用例上下文
- 仅说「PRD 转文档」而未提及测试用例或目标格式
- 仅讨论测试策略或测试方法，不涉及文件生成
- 生成非测试用例类的 XMind/Excel（如项目计划、数据分析）

---

## 二、输出格式选择

用户未指定格式时，默认输出 XMind。可通过用户指令显式选择：

| 用户指令 | 输出格式 |
|---------|---------|
| 「生成 XMind」 | .xmind 文件 |
| 「生成 Excel」 | .xlsx 文件 |
| 「两种都要」 | 同时生成 .xmind 和 .xlsx |

---

## 三、输入处理

### 3.1 PRD / 需求文档

1. **获取原始文档**：用户直接粘贴 PRD/需求/功能说明
2. **识别文档类型**：业务需求、功能说明或混合型
3. **提取关键信息**：
   - 功能模块与子模块
   - 核心业务流程与用户场景
   - 输入参数与约束条件
   - 角色与权限
   - 状态流转
   - 性能指标（非必需）

### 3.2 接口文档

1. **获取原始文档**：用户粘贴接口说明、Swagger JSON/YAML、Postman Collection 等
2. **识别文档格式**：
   - OpenAPI/Swagger 格式（JSON/YAML）
   - Markdown 格式的接口说明
   - 自定义格式的接口列表
3. **提取关键信息**：
   - 接口路径与方法（GET/POST/PUT/DELETE）
   - 请求参数（header/query/body/path）
   - 响应状态码与响应体结构
   - 认证方式（Bearer Token/API Key/OAuth）
   - 错误码定义

### 3.3 Excel 测试用例文件

1. **获取文件**：用户上传 .xlsx/.xls 文件
2. **解析 Excel**：使用 Python + openpyxl 读取
3. **识别列结构**（按优先级匹配）：
   - 标准格式：`用例编号 | 模块 | 用例名称 | 优先级 | 前置条件 | 测试步骤 | 预期结果`
   - 简化格式：`模块 | 用例名称 | 步骤 | 预期`
   - 自定义格式：根据表头识别关键列
4. **映射到目标格式**（XMind 或 Excel）

**Excel 解析命令**：
```bash
# 安装依赖
pip install openpyxl

# 解析 Excel 文件
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('用例文件.xlsx')
for sheet in wb.sheetnames:
    ws = wb[sheet]
    headers = [cell.value for cell in ws[1]]
    print(f'Sheet: {sheet}, Headers: {headers}')
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:  # 跳过空行
            print(row)
"
```

---

## 四、测试用例设计策略（默认运用）

### 4.1 正向测试用例（正常流程）
- 每个核心功能至少1条正向用例，验证正常路径

### 4.2 反向/异常测试用例（负向用例）
- 非法输入、错误操作、异常状态，验证系统正确拒绝

### 4.3 边界值用例
- 最小值、最大值、刚好超界、空值、长度临界点

### 4.4 等价类划分
- 有效等价类与无效等价类，每类选取代表值

### 4.5 状态与流程策略
- 合法状态迁移与非法迁移，角色切换与权限覆盖

### 4.6 场景法/用户场景
- 端到端用户故事，跨模块串联用例

---

## 五、输出格式

### 5.1 XMind 格式

#### 5.1.1 XMind 8 ZIP 结构

XMind 文件本质是 ZIP 包，包含以下3个文件（**必须**）：

```
xxx.xmind/
├── content.json    # 思维导图内容（XMind专用JSON格式）
├── metadata.json   # 元数据（必须包含版本信息）
└── manifest.json   # 文件清单（必须包含file-entries）
```

> **⚠️ 重要**：content.json 不是通用JSON格式，必须使用 XMind 专用结构，否则 XMind 会提示 "not a valid XMind File"。

#### 5.1.2 content.json 格式（XMind专用结构）

```json
[
  {
    "id": "全局唯一ID",
    "revisionId": "修订版本ID",
    "class": "sheet",
    "rootTopic": {
      "id": "根主题ID",
      "class": "topic",
      "title": "【项目名】测试用例",
      "children": {
        "attached": [
          {
            "id": "模块ID",
            "class": "topic",
            "title": "1. 功能模块名",
            "children": {
              "attached": [
                {
                  "id": "子模块ID",
                  "class": "topic",
                  "title": "1.1 子功能点",
                  "children": {
                    "attached": [
                      {
                        "id": "用例ID",
                        "class": "topic",
                        "title": "【正向】用例名称-P0",
                        "notes": {
                          "plain": {
                            "content": "前置条件：xxx\n步骤：1.xxx 2.xxx\n预期：xxx"
                          }
                        },
                        "markers": [
                          {
                            "markerId": "priority-1"
                          }
                        ]
                      }
                    ]
                  }
                }
              ]
            }
          }
        ]
      }
    }
  }
]
```

**关键结构说明**：
- 顶层是**数组** `[]`，不是对象 `{}`
- `class: "sheet"` 表示工作表
- `class: "topic"` 表示主题节点
- 子节点在 `children.attached` 数组中（不是 `children` 数组）
- `notes.plain.content` 是备注内容（支持 `\n` 换行）
- `markers.markerId` 是优先级标记：`priority-1`（P0/红星）、`priority-2`（P1/橙星）、`priority-3`（P2/黄星）

#### 5.1.3 metadata.json 格式

```json
{
  "dataStructureVersion": "2",
  "layoutEngineVersion": "3",
  "creator": {
    "name": "Claude",
    "version": "1.0"
  }
}
```

#### 5.1.4 manifest.json 格式

```json
{
  "file-entries": {
    "content.json": {},
    "metadata.json": {}
  }
}
```

### 5.2 Excel 格式

#### 5.2.1 Excel 标准格式

```
| 用例编号 | 所属模块 | 用例标题 | 优先级 | 用例类型 | 前置条件 | 测试步骤 | 预期结果 |
|---------|---------|---------|--------|---------|---------|---------|---------|
| TC-001  | 认证模块 | 管理员登录成功 | P0 | 功能测试 | 系统中存在管理员账号admin | 1. POST /api/auth/login | 1. HTTP 200 2. code=200 |
| TC-002  | 认证模块 | 用户名密码错误 | P1 | 功能测试 | - | 1. POST /api/auth/login | 1. HTTP 200 2. code=401 |
```

**模板文件**：`docs/testcases_template.xlsx`

#### 5.2.2 Excel 列说明

| 列名 | 必填 | 说明 | 示例 |
|-----|------|------|------|
| 用例编号 | 是 | 唯一标识，格式 TC-XXX | TC-001 |
| 所属模块 | 是 | 功能模块名称 | 认证模块、用户模块 |
| 用例标题 | 是 | 清晰表达验证点 | 管理员登录成功 |
| 优先级 | 是 | P0/P1/P2/P3 | P0 |
| 用例类型 | 是 | 功能测试/边界测试/异常测试等 | 功能测试 |
| 前置条件 | 否 | 执行前提条件 | 已登录获取token |
| 测试步骤 | 是 | 1.2.3.步骤描述，支持多行 | 1. POST /api/users |
| 预期结果 | 是 | 预期输出/响应，支持多行 | 1. HTTP 200 2. code=200 |

#### 5.2.3 Excel 生成代码模板

```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_testcases_excel(data, output_path):
    """
    data: List of [用例编号, 所属模块, 用例标题, 优先级, 用例类型, 前置条件, 测试步骤, 预期结果]
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "接口测试用例"

    # 样式定义
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    title_font = Font(bold=True, size=14)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # 优先级颜色填充（浅色背景，便于阅读）
    priority_fills = {
        "P0": PatternFill(start_color="FFE0E0", end_color="FFE0E0", fill_type="solid"),  # 浅红色
        "P1": PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid"),  # 浅橙色
        "P2": PatternFill(start_color="FFFDE7", end_color="FFFDE7", fill_type="solid"),  # 浅黄色
        "P3": PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid"),  # 浅绿色
    }
    priority_font_colors = {
        "P0": "C62828",  # 深红色
        "P1": "E65100",  # 深橙色
        "P2": "F9A825",  # 深黄色
        "P3": "2E7D32",  # 深绿色
    }

    # 第1行：标题
    ws.merge_cells('A1:H1')
    ws.cell(row=1, column=1, value="📋 接口测试用例")
    ws.cell(row=1, column=1).font = title_font
    ws.cell(row=1, column=1).alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 25

    # 第2行：统计信息（可选，生成时动态填充）
    # ws.cell(row=2, column=1, value="总计: N 条 | P0: X | P1: Y | P2: Z | P3: W")

    # 第3行：表头
    headers = ["用例编号", "所属模块", "用例标题", "优先级", "用例类型", "前置条件", "测试步骤", "预期结果"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    ws.row_dimensions[3].height = 20

    # 数据行
    for row_idx, row_data in enumerate(data, start=4):
        priority = row_data[3] if len(row_data) > 3 else "P2"  # 默认P2
        priority_fill = priority_fills.get(priority)
        priority_font_color = priority_font_colors.get(priority, "000000")

        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = thin_border
            cell.alignment = Alignment(vertical='top', wrap_text=True)

            # 优先级列（第4列）添加颜色背景和字体颜色
            if col_idx == 4 and priority_fill:
                cell.fill = priority_fill
                cell.font = Font(bold=True, color=priority_font_color)

    # 添加筛选器（表头行 + 数据区域）
    ws.auto_filter.ref = f"A3:H{3 + len(data)}"

    # 列宽设置
    col_widths = {'A': 12, 'B': 12, 'C': 25, 'D': 8, 'E': 12, 'F': 25, 'G': 40, 'H': 40}
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # 冻结首行（标题行）和表头行
    ws.freeze_panes = 'A4'

    wb.save(output_path)
    return output_path

# 使用示例
if __name__ == "__main__":
    testcases = [
        ["TC-001", "认证模块", "管理员登录成功", "P0", "功能测试",
         "系统中存在管理员账号admin，密码为123456",
         "1. 使用POST方法访问/api/auth/login接口\n2. 请求体中设置username为admin，password为123456",
         "1. 返回HTTP状态码200\n2. 响应体中code字段为200\n3. data.token字段非空"],
        ["TC-002", "认证模块", "普通用户登录成功", "P0", "功能测试",
         "系统中存在普通用户账号reader01，密码为reader123",
         "1. 使用POST方法访问/api/auth/login接口\n2. 请求体中设置username为reader01，password为reader123",
         "1. 返回HTTP状态码200\n2. 响应体中code字段为200\n3. data.token字段非空"],
        ["TC-003", "认证模块", "密码错误登录失败", "P1", "功能测试",
         "-",
         "1. POST /api/auth/login with wrong password",
         "1. HTTP 401\n2. code=401"],
        ["TC-004", "用户模块", "查询用户列表", "P2", "功能测试",
         "已获取有效token",
         "1. GET /api/users with Authorization header",
         "1. HTTP 200\n2. 返回用户列表"],
    ]
    create_testcases_excel(testcases, "testcases_项目名_20260504.xlsx")
```

---

## 六、工作流

### 6.1 PRD / 需求文档 → XMind / Excel

1. **解析文档**：提取模块、功能点、业务流程、约束条件
2. **设计用例**：运用6大测试策略生成用例
3. **生成目标文件**（根据用户选择）
4. **输出文件路径**

### 6.2 接口文档 → XMind / Excel

1. **解析接口定义**：提取接口路径、方法、参数、响应
2. **设计用例**：运用 API 测试策略
   - 参数校验用例
   - 鉴权/授权用例
   - 错误码覆盖用例
   - 业务逻辑用例
3. **生成目标文件**
4. **输出文件路径**

### 6.3 Excel → XMind

1. **读取 Excel 文件**：解析工作表和列结构
2. **映射数据**：
   - Sheet 名 → 项目名
   - 模块列 → 一级节点
   - 用例名列 → 子节点
   - 优先级 → markers
   - 步骤/预期 → notes
3. **生成 XMind 文件**
4. **输出文件路径**

### 6.4 XMind → Excel

1. **读取 XMind 文件**：解压并解析 content.json
2. **映射数据**：
   - 一级节点 → 模块
   - 子节点 → 用例名称
   - notes.content → 解析为步骤/预期
   - markers → 优先级
3. **生成 Excel 文件**
4. **输出文件路径**

### 6.5 通用 XMind 生成步骤

```bash
# 创建临时目录
mkdir -p temp_xmind

# 写入 content.json（XMind专用JSON格式 - 数组结构）
# 写入 metadata.json
# 写入 manifest.json

# 打包成 xmind
# Windows PowerShell
Compress-Archive -Path 'content.json','metadata.json','manifest.json' -DestinationPath 'testcases_项目名.zip' -Force
Move-Item testcases_项目名.zip testcases_项目名_YYYYMMDD.xmind

# Unix/macOS
cd temp_xmind && zip -r ../testcases_项目名_YYYYMMDD.xmind * && cd ..

# 清理临时目录
rm -rf temp_xmind
```

---

## 七、标记说明

### 7.1 XMind 标记

| 标记 | 含义 | XMind Maker |
|------|------|-------------|
| P0 | 核心用例，必须通过 | priority-1 (红星) |
| P1 | 重要用例，建议通过 | priority-2 (橙星) |
| P2 | 普通用例，常规覆盖 | priority-3 (黄星) |
| 【正向】 | 正常流程用例 | - |
| 【反向】 | 异常/负向用例 | - |
| 【边界】 | 边界值用例 | - |
| 【等价】 | 等价类用例 | - |
| 【状态】 | 状态流转用例 | - |
| 【场景】 | 端到端场景用例 | - |
| 【API】 | 接口测试用例 | - |

### 7.2 Excel 类型列标记

| 类型 | 含义 |
|------|------|
| 正向 | 正常流程用例 |
| 反向 | 异常/负向用例 |
| 边界 | 边界值用例 |
| 等价 | 等价类用例 |
| 状态 | 状态流转用例 |
| 场景 | 端到端场景用例 |

---

## 八、用例数量估算参考

### 8.1 PRD / 需求文档

| 模块复杂度 | 正向用例 | 反向用例 | 边界用例 | 状态/场景用例 |
|-----------|---------|---------|---------|--------------|
| 简单（1-3个功能点） | 3-5 | 5-8 | 3-5 | 2-3 |
| 中等（4-7个功能点） | 6-10 | 10-15 | 5-10 | 5-8 |
| 复杂（8+功能点） | 12-20 | 20-30 | 10-15 | 10-15 |

### 8.2 接口文档

| 接口复杂度 | 参数校验用例 | 鉴权用例 | 错误码用例 | 业务逻辑用例 |
|-----------|------------|---------|-----------|-------------|
| 简单（1-5个接口） | 10-20 | 5-10 | 5-10 | 5-10 |
| 中等（6-15个接口） | 20-50 | 10-20 | 10-20 | 10-20 |
| 复杂（16+个接口） | 50-100 | 20-40 | 20-40 | 20-40 |

### 8.3 Excel 转换

- Excel 中有多少条用例，XMind 就生成多少条（反之亦然）
- 按原有的模块分组组织

---

## 九、结合 references 标准文档

根据文档类型，可叠加 `references/` 下的专用标准增强覆盖：

| 文档类型 | 参考标准 | 增强内容 |
|---------|---------|---------|
| 功能需求 | `references/functional-testcases-standard.md` | 业务流程、状态流转、角色权限 |
| 接口文档 | `references/api-testcases-standard.md` | 请求参数校验、错误码、鉴权、幂等 |
| 性能需求 | `references/performance-testcases-standard.md` | RT、QPS、吞吐量、稳态/峰值 |
| 自动化候选 | `references/automation-testcases-standard.md` | 标记自动化适用性 |

**优先级**：SKILL 通用策略（正向/反向/边界等）为基础 → 按文档类型叠加 references 标准。

---

## 十、质量自检

### 10.1 通用检查
- [ ] 用例是否按模块层级清晰组织？
- [ ] 每个核心功能是否至少1条正向用例？
- [ ] 用例标题是否清晰表达验证点？

### 10.2 PRD / 需求文档检查
- [ ] 关键约束是否都有反向/异常覆盖？
- [ ] 有范围/长度限制的是否有边界值用例？
- [ ] 多角色/多状态是否有权限与状态流转用例？
- [ ] 典型用户场景是否覆盖端到端路径？

### 10.3 接口文档检查
- [ ] 每个接口是否都有正向用例？
- [ ] 必填参数缺失是否都有覆盖？
- [ ] 参数类型/格式/长度校验是否完整？
- [ ] 所有错误码是否可触发？
- [ ] 鉴权失败场景是否覆盖？
- [ ] 权限不足场景是否覆盖？

### 10.4 格式转换检查
- [ ] XMind/Excel 转换是否保持数据一致性？
- [ ] 用例编号是否保持唯一？
- [ ] 模块分组是否正确？
- [ ] 优先级和类型标记是否一致？

---

## 十一、保存与落盘

- **默认输出**：根据用户选择输出 `.xmind` 或 `.xlsx`
- **保存位置**：当前工作目录
- **文件名格式**：

| 输出类型 | 文件名格式 |
|---------|-----------|
| PRD/需求 → XMind | `testcases_【项目名】_YYYYMMDD.xmind` |
| PRD/需求 → Excel | `testcases_【项目名】_YYYYMMDD.xlsx` |
| 接口文档 → XMind | `testcases_【项目名】_API_YYYYMMDD.xmind` |
| 接口文档 → Excel | `testcases_【项目名】_API_YYYYMMDD.xlsx` |
| Excel → XMind | `testcases_【原文件名】_YYYYMMDD.xmind` |
| XMind → Excel | `testcases_【原文件名】_YYYYMMDD.xlsx` |

- **生成命令（Windows PowerShell）**：
  ```powershell
  # XMind 生成
  mkdir -p temp_xmind
  # 写入 content.json、metadata.json、manifest.json
  Compress-Archive -Path 'content.json','metadata.json','manifest.json' -DestinationPath 'testcases.zip' -Force
  Move-Item testcases.zip testcases_项目名_YYYYMMDD.xmind
  rm -rf temp_xmind

  # Excel 生成
  # 使用 Python + openpyxl 生成
  ```

- **生成命令（Unix/Linux/macOS）**：
  ```bash
  # XMind 生成
  mkdir -p temp_xmind
  # 写入 content.json、metadata.json、manifest.json
  cd temp_xmind && zip -r ../testcases_项目名_YYYYMMDD.xmind * && cd ..
  rm -rf temp_xmind

  # Excel 生成
  # 使用 Python + openpyxl 生成
  ```

> **⚠️ XMind 常见错误**：`content.json` 使用通用JSON格式（对象结构 `{}`）而非XMind专用格式（数组结构 `[]`），导致 XMind 提示 "not a valid XMind File"。
