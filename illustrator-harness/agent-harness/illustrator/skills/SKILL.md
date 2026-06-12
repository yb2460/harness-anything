---
name: "cli-anything-illustrator"
description: "Illustrator CLI — 通过命令行操控 Adobe Illustrator 矢量图形"
---

# cli-anything-illustrator

通过 Windows COM 自动化接口操控 Adobe Illustrator 的命令行工具。支持文档管理、图层操作、形状绘制、文字编辑和格式导出。

## 前置条件

- **Windows 操作系统**
- **Adobe Illustrator 2023+**
- **Python 3.10+**
- **pywin32**：`pip install pywin32`
- **click**：`pip install click`

## 安装

```bash
cd illustrator-harness/agent-harness
pip install -e .
```

## 命令结构

```
illustrator
├── project          # 文档管理
│   ├── new          # 新建 AI 文档
│   ├── open         # 打开文档
│   └── save         # 保存文档
├── layers           # 图层管理
│   ├── add          # 添加图层
│   ├── remove       # 删除图层
│   ├── rename       # 重命名图层
│   ├── show         # 显示图层
│   └── hide         # 隐藏图层
├── shapes           # 形状绘制
│   ├── rect         # 矩形
│   ├── ellipse      # 椭圆
│   ├── line         # 线条
│   └── polygon      # 多边形
├── text             # 文字编辑
│   ├── add          # 添加文字
│   └── modify       # 修改文字属性
└── export           # 导出
    ├── png          # 导出 PNG
    ├── jpeg         # 导出 JPEG
    ├── svg          # 导出 SVG
    ├── pdf          # 导出 PDF
    └── ai           # 导出 AI
```

## 工作原理

通过 `illustrator_cli.py` → `ai_backend.py` → COM Bridge → `Illustrator.Application` 实现命令行到 Illustrator 引擎的完整调用链。
