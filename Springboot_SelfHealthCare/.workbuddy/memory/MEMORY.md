# 项目长期记忆

## 项目概述
- 项目：基于SpringBoot的个人健康管理与疾病预警系统（毕业设计）
- 技术栈：Spring Boot 3.5.11 / Java 21 / Spring Data JPA / H2文件数据库（MySQL兼容模式）/ Session认证（BCrypt）/ Apache PDFBox 3.0.5 / Tesseract OCR / DeepSeek Chat API
- 前端：原生 HTML + CSS + JavaScript（无框架）
- 端口：8081
- 启动脚本：`start-app.bat`

## 系统功能模块（8大模块，20个API端点）
1. 用户认证 `/api/auth`（4端点）：注册/登录/退出/会话查询，BCrypt+HttpSession
2. 健康总览仪表盘 `/api/dashboard`（1端点）：统计摘要+最近记录/预警
3. 个人健康档案 `/api/profile`（3端点）：档案CRUD（upsert），24个字段
4. 健康记录 `/api/records`（5端点）：22项指标 CRUD+分页+风险筛选
5. 风险评估与预警 `/api/alerts`（4端点）：自动评估+预警生成/状态更新
6. 数据可视化 `/api/visualization`（1端点）：近20条记录6指标趋势系列
7. 健康文档导入 `/api/imports`（1端点）：PDF/OCR/文本三类格式，26字段正则解析
8. AI咨询 `/api/ai`（1端点）：DeepSeek Chat API集成，需DEEPSEEK_API_KEY环境变量

## 核心实体关系
- AppUser (1) → UserProfile (1:1)
- AppUser (1) → HealthRecord (1:N)
- HealthRecord (1) → HealthAlert (1:N)
- BaseEntity 提供 id/createdAt/updatedAt 公共字段

## 关键业务链路
- 健康记录保存 → RiskAssessmentService.assess()（12类指标评估）→ 设置bmi/riskScore/riskLevel/summary → 保存记录 → refreshAlerts()（先删后增）

## 重要文件
- `update_doc.py`：毕业设计手册更新脚本（已执行，2026-03-30）
- `基于SpringBoot的个人健康管理与疾病预警系统（更新版）.docx`：更新完成的毕业设计手册
- `部署环境配置清单.txt`：部署环境说明
