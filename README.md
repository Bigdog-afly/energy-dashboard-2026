# 鄂尔多斯供电公司 · 2026年十项民生工程进度看板

> 交互式 BI 数据看板，实时跟踪十项民生工程治理进度

## 📊 在线访问

- **看板地址**: https://bigdog-afly.github.io/energy-dashboard-2026/
- **GitHub 仓库**: https://github.com/Bigdog-afly/energy-dashboard-2026

## 🏗️ 项目结构

```
├── dashboard.html          # 原始看板（自定义图表）
├── dashboard_v2.html       # 增强版看板（ECharts 可视化）⭐ 推荐
├── dashboard_data.json     # 看板数据源
├── build_data.js           # Excel → JSON 数据构建脚本
├── upload_report.js        # 周报上传与数据处理脚本
├── package.json            # Node.js 依赖配置
├── report_history/         # 历史数据快照目录
│   ├── *.json              # 时间戳命名的数据快照
│   └── changes.md          # 变更日志
└── .github/workflows/
    └── update-dashboard.yml # GitHub Actions 自动化配置
```

## 🚀 快速开始

### 本地查看看板

1. 确保已安装 Node.js
2. 安装依赖：
   ```bash
   npm install
   ```
3. 启动本地服务器：
   ```bash
   npx http-server . -p 8080
   ```
4. 浏览器打开 http://localhost:8080/dashboard_v2.html

> 💡 由于浏览器的 CORS 限制，看板数据需要通过 HTTP 服务器加载，不能直接双击 HTML 文件打开。

### 更新看板数据

#### 方法一：使用 upload_report.js（推荐）

```bash
# 处理默认文件
node upload_report.js

# 指定文件路径
node upload_report.js "path/to/new-report.xlsx"

# 预览模式（不写入文件）
node upload_report.js --dry-run
```

此脚本会：
- 解析 Excel 文件中的所有工程数据
- 生成 `dashboard_data.json`
- 保存历史快照到 `report_history/`
- 输出变更摘要和趋势分析

#### 方法二：使用 build_data.js

```bash
node build_data.js
```

#### 方法三：通过 GitHub Actions 自动更新

1. 将新的 Excel 文件推送到仓库
2. 或在仓库页面点击 **Actions** → **更新民生工程进度看板数据** → **Run workflow**
3. 可选：提供 Excel 文件下载 URL

## 📈 看板功能

### 可视化图表（ECharts v2）

| 图表 | 说明 |
|------|------|
| 🥧 饼图 | 项目状态分布（已完成/进行中/未开始） |
| 🕸️ 雷达图 | 十项工程综合完成率一览 |
| 📊 柱状图 | 各工程投资对比 |
| 📈 进度条 | 各工程完成率排名 |
| 📋 明细表 | 工程维度详细数据 |

### 数据模块

1. **总体情况概览** - 总项目数、已完成、总投资、整体完成率
2. **项目状态分布** - 饼图展示三类状态占比
3. **十项工程完成率** - 雷达图 + TOP5 排行
4. **预警中心** - 自动识别滞后工程并发出预警
5. **督办事项** - 完成率低于 30% 的工程重点标注
6. **示范区专项** - 6 大示范区详细数据
7. **农牧区线路专项** - 各地市线路治理进度

## 🔧 技术栈

- **前端**: 纯 HTML/CSS/JavaScript（零框架）
- **可视化**: Apache ECharts 5.x
- **数据处理**: Node.js + SheetJS (xlsx)
- **部署**: GitHub Pages（免费静态托管）
- **CI/CD**: GitHub Actions（定时自动更新）

## 📝 数据字段说明

### sections.{工程名}
| 字段 | 说明 |
|------|------|
| t | 总项目数 |
| d / tot | 已完成数 |
| inv | 进行中数 |
| nst | 未开始数 |
| pct | 完成率(%) |
| totalInvest | 总投资(万元) |
| bc | 按地市/分公司分组数据 |

### grand
| 字段 | 说明 |
|------|------|
| total | 所有工程总项目数 |
| done | 已完成总数 |
| investing | 进行中总数 |
| notstart | 未开始总数 |
| invest | 总投资(万元) |
| pct | 整体完成率(%) |

## 🔒 安全提示

- Token 建议定期轮换
- 不要在代码中硬编码敏感信息
- GitHub Actions 使用 GITHUB_TOKEN 自动认证

---

鄂尔多斯供电公司 · 生产技术部 · 2026
