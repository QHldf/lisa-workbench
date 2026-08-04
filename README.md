# Lisa 融资顾问工作台 · 完整技术文档

> **用途**：将本目录完整复制到任意智能体环境，可 100% 复现完整功能。
> **环境要求**：Python 3.8+, Flask, python-docx。浏览器端无需任何依赖。

---

## 零、快速启动（3步）

```bash
# 步骤1：创建 templates 目录（index.html 需要放在这里）
mkdir -p templates
mv index.html templates/

# 步骤2：安装依赖
pip install flask python-docx

# 步骤3：启动
python app.py
# 默认监听 http://0.0.0.0:8000
# 可通过 PORT=5000 python app.py 修改端口
```

**启动后目录结构应为：**
```
./
  ├── app.py                          # Flask 后端（1043行）
  ├── templates/
  │     └── index.html                # 单页前端（889行，零依赖）
  └── knowledge-base/
        └── bank-data/
              ├── workbench_data.json # 95客户 + 75银行 + 14在做进展
              ├── products.json       # 246款银行产品政策
              └── rules.json          # 分层/获贷率/红线规则
```

---

## 一、数据文件结构

### 1.1 workbench_data.json（核心数据）

```json
{
  "clients": [{
    "name": "杭州XX科技有限公司",           // 客户名称（主键）
    "tag": "国高新、科技型中小企业",          // 企业标签
    "income": 1600.0,                        // 年开票（万）
    "tax": 15.0,                             // 年纳税（万）
    "biz_inst_count": 3,                     // 企业端机构数
    "bank_count": 5,                         // 在贷银行数
    "all_institutions": ["工商银行","建设银行"], // 在贷银行名单
    "template": "客户档案：杭州XX...",         // 完整档案纯文本（单一真相源）
    "created": "2025-08-04",                 // 创建日期
    "legal_age": 49,                         // 法人年龄
    "established_year": 2021,                // 成立年份
    "has_overdue": false,                    // 是否有逾期
    "overdue_cont": 0,                       // 连三次数
    "overdue_cum": 0,                        // 累六次数
    "debt_total": 400.0,                     // 企业总负债（万）
    "is_dual_sign": true,                    // 是否双签
    "is_gaoxin": true,                       // 是否国高新
    "has_tech": true,                        // 是否有科技标签
    "query_6m": 8, query_12m": 15,           // 查询次数
    "cc_usage": 35.0,                        // 信用卡使用率%
    "tags": "国高新、科技型中小企业",          // 企查查验证后标签
    "enterprise_labels": "国高新、科技型中小企业"
  }],
  "insts": [{ "name": "工商银行", "count": 15 }],
  "active": [{
    "name": "杭州XX", "status": "推进中",     // 紧急/推进中/待反馈/暂停
    "deadline": "2025-08-10", "overdue": false,
    "progress": "已提交材料给工行...", "next": "周三前拿到审批结果"
  }]
}
```

### 1.2 products.json（银行产品，每条记录格式）

```json
{
  "bank": "建设银行", "product": "云税贷（纯信用）",
  "type": "纯信用", "amount_max": 300, "rate": "3.40%-4.55%",
  "rate_min": 3.0, "rate_max": 4.55,
  "min_age": 18, "max_age": 65, "tax_grades": ["A","B","M"],
  "min_tax": 0.5, "min_established": 2,
  "query_3m_max": 6, "query_6m_max": 12, "inst_max": 2,
  "debt_max": 500, "overdue_rule": "无连3累6",
  "min_shareholding": 50, "dual_sign_required": true
}
```

### 1.3 rules.json

```json
{
  "tiers": { "工商银行": "T1", "中国银行": "T1", "建设银行": "T2", ... },
  "rates": { "工商银行": "39.5%", "中国银行": "27.6%", ... },
  "redline": [],
  "tier_bonus": { "T1": 15, "T2": 8, "T3": 3, "T4": 1 }
}
```

---

## 二、核心 API 接口

| 路由 | 方法 | 说明 |
|------|------|------|
| `/` | GET | 渲染工作台首页 |
| `/api/stats` | GET | 统计（客户数/银行数/产品数/在做/紧急/逾期） |
| `/api/active` | GET | 在做客户进展（按紧急度+截止日期排序） |
| `/api/clients?q=&tag=&bank=&active=` | GET | 客户列表（搜索/标签/银行/在做过滤） |
| `/api/client/<name>` | GET | 客户详情（含在做进展） |
| `/api/client/add` | POST | 新增/更新客户（模板为单一真相源） |
| `/api/banks?tier=` | GET | 银行列表（分层过滤） |
| `/api/bank/<name>` | GET | 银行详情（含产品列表） |
| `/api/bank/add` | POST | 新增/更新银行产品 |
| `/api/rules` | GET | 读取规则全量 |
| `/api/rules/update` | POST | 更新分层/获贷率/红线 |
| `/api/active/update` | POST | 更新在做进展 |
| `/api/reload` | POST | 手动全量重载（改JSON后调用） |
| `/api/backfill` | POST | 全量同步：模板→结构化字段 |
| `/api/tags` | GET | 全部标签列表 |
| `/api/all_banks` | GET | 全部银行名称列表 |
| `/api/match` | POST | 智能匹配引擎 |
| `/api/report/<name>` | GET | 导出诊断报告 DOCX |

---

## 三、核心机制

### 3.1 单一真相源
- 客户档案模板纯文本是**唯一真相源**
- `sync_client_from_template()` 从模板解析 income/tax/biz_inst_count/在贷银行等
- `/api/backfill` 可全量重同步

### 3.2 实时同步
- `@app.before_request` 每次请求前从磁盘重载 JSON
- 改 JSON = 即时生效，无需重启

### 3.3 智能匹配引擎
```
输入：年龄/开票/纳税/纳税级别/查询次数/机构数/负债/信用卡/销贷比/逾期/科技标签/排除银行
  ↓
1. 风控预检：连三累六一票否决 / 高风险项告警
2. 遍历所有银行产品逐条匹配
3. 按匹配分排序（tier_bonus + 产品匹配 + 获贷率）
4. 排除存量/征信/已还清银行
  ↓
输出：Top5推荐 + 完整匹配列表 + 匹配/排除理由
```

### 3.4 模板解析规则

`parse_template()` 正则提取：
- 法人年龄：`（49岁）` / `49岁`
- 成立年份：`2021年成立`、持股：`持股100%`
- 纳税级别：`纳税A级`
- 查询：`近3月查询X次` / `近6个月查询X次`
- 信用卡：`使用率35%`、销贷比：`销贷比40%`
- 逾期：`征信逾期：无` / `连3累6`
- 总负债：`借贷余额共计XX万` / `合计约XX万`
- 双签：检测"双签"关键词

`extract_income_from_tpl()` 优先级：
1. 整年开票（`2023年开票1600万`）→ 取最新年份
2. 非整年开票 → 取最后一个
3. 应税销售额/营收/流水/营业额 → 兜底

`extract_institutions_from_tpl()`：
- 只从「企业负债」段落提取
- 简称→全名映射（工行→工商银行）
- 去噪：排除"银行承兑汇票"等非机构名

### 3.5 银行简称映射

```python
BANK_ALIASES = {
    '工行': '工商银行', '建行': '建设银行', '农行': '农业银行', '中行': '中国银行',
    '交行': '交通银行', '邮储': '邮储银行', '招行': '招商银行', '中信': '中信银行',
    '浦发': '浦发银行', '民生': '民生银行', '兴业': '兴业银行', '光大': '光大银行',
    '华夏': '华夏银行', '平安': '平安银行', '浙商': '浙商银行', '渤海': '渤海银行',
    '广发': '广发银行', '恒丰': '恒丰银行',
    '新网': '新网银行', '微众': '微众银行', '网商': '网商银行',
    '宁波': '宁波银行', '杭州': '杭州银行', '南京': '南京银行',
    '江苏': '江苏银行', '北京': '北京银行', '上海': '上海银行',
    '杭州联合': '杭州联合农商行', '余杭农商': '余杭农商行', '萧山农商': '萧山农商行',
    '临安农商': '临安农商行', '泰隆': '泰隆银行', '民泰': '民泰银行',
    '温州': '温州银行', '稠州': '稠州银行',
}
```

---

## 四、前端功能

| 模块 | 功能 |
|------|------|
| 今日作战区 | 在做客户排序、逾期红色高亮、截止倒计时 |
| 客户列表 | 搜索+标签/银行/在做三联动筛选、点击看详情 |
| 客户详情 | 画像速览、排除清单、一键智能匹配、导出DOCX |
| 银行列表 | T1-T4分层+进度条、获贷率/客户数/产品数 |
| 智能匹配 | 17字段表单、风控预检、Top5推荐 |
| 管理端 | 分层/获贷率/红线编辑、产品政策编辑、进展更新 |

---

## 五、部署到任意智能体

```bash
# 本目录已包含所有文件。直接：
mkdir -p templates
mv index.html templates/
pip install flask python-docx
python app.py

# 端口默认 8000，可通过 PORT=5000 python app.py 修改
# 数据路径自动兼容 /workspace/knowledge-base 和 ../knowledge-base
# 所有改动通过 API 写盘，无需重启
# 前端纯 HTML/CSS/JS，无框架依赖
```

---

## 六、文件清单

| 文件 | 大小 | 行数 | 说明 |
|------|------|------|------|
| app.py | 42KB | 1043 | Flask 后端（全部业务逻辑） |
| templates/index.html | 59KB | 889 | 单页前端（5Tab+弹窗+匹配） |
| knowledge-base/bank-data/workbench_data.json | 218KB | — | 95客户+75银行+14在做 |
| knowledge-base/bank-data/products.json | 89KB | — | 246款银行产品政策 |
| knowledge-base/bank-data/rules.json | 1.3KB | — | 分层/获贷率/红线 |

---

> **所有源码均在本目录中，无需额外查找。**
