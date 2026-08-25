# 智伴成长（ZhiBan-ChengZhang）

**“智伴成长”——基于 AI 与朋辈导航的大学生全周期智能导学系统**

大学生创新创业训练计划项目（创新训练项目），2026.09—2027.06

> 让每一段大学时光，都有智慧与温暖相伴。

## 项目简介

当代大学生面临信息检索困境与校园适应断层双重挑战。本项目以 RAG（检索增强生成）与知识图谱（Neo4j）为核心技术，构建面向高校场景的“数字学长”智能导学系统，覆盖生活适应、学业规划、心理疏导、校园融入四大维度，采用 B2B 订阅制商业模式（学校付费、学生免费）。

## 文件清单

| 文件 | 说明 |
|---|---|
| `智伴成长_项目计划书.docx` | 大创项目计划书（申报材料） |
| `智伴成长_商业计划书.docx` | 商业计划书终稿（市场/产品/商业模式/财务预测） |
| `智伴成长_答辩PPT.pptx` | 答辩演示文稿（19 页，16:9） |
| `智伴成长.docx` | 开题报告 |
| `智伴成长_logo_icon.png` | 方形图标（1024×1024） |
| `智伴成长_logo_横版.png` | 横版 Logo（透明底） |
| `智伴成长_logo_横版.jpg` | 横版 Logo（白底） |
| `gen_plan.py` | 生成项目计划书的脚本 |
| `gen_business_plan.py` | 生成商业计划书的脚本 |
| `gen_defense_ppt.py` | 生成答辩 PPT 的脚本 |
| `gen_logo.py` | 生成 Logo 的脚本 |

## 重新生成

所有文档均由脚本生成，修改脚本后重跑即可：

```bash
python3 gen_plan.py            # 项目计划书
python3 gen_business_plan.py   # 商业计划书（Word 打开后 Ctrl+A + F9 更新目录）
python3 gen_defense_ppt.py     # 答辩 PPT
python3 gen_logo.py            # Logo 三件套
```

依赖：`pip install python-docx python-pptx pillow`
