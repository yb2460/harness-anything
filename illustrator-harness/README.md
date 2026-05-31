# Illustrator Harness

> CLI-Anything Illustrator —— 通过 COM 自动化操控 Adobe Illustrator 的 AI Agent 工具集。

## 概述

Illustrator Harness 是 CLI-Anything 生态的 Illustrator 适配器，让 AI Agent 可以通过**命令行 + COM 自动化接口**直接操控 Adobe Illustrator，实现矢量图形编辑、设计稿生成、批量处理等任务的自动化。

## 目录结构

```
illustrator-harness/
└── agent-harness/
    ├── setup.py                     # pip 安装配置
    ├── test_ai_harness.py           # AI 集成测试
    └── illustrator/
        ├── __init__.py
        ├── __main__.py
        ├── illustrator_cli.py       # CLI 入口（Click 框架）
        ├── core/                    # 核心模块
        │   ├── __init__.py
        │   ├── project.py           #   项目管理（新建/打开/保存）
        │   ├── layers.py            #   图层操作
        │   ├── shapes.py            #   形状绘制（矩形/椭圆/线条）
        │   ├── text.py              #   文字操作（字体/大小/颜色）
        │   └── export.py            #   导出（PNG/JPEG/SVG/PDF/AI）
        ├── utils/
        │   ├── __init__.py
        │   ├── ai_backend.py        #   COM 桥接层（Illustrator.Application）
        │   └── repl_skin.py         #   交互式 REPL 界面
        └── tests/
            └── __init__.py
```

## 架构设计

```
CLI 命令 → Click CLI 层 → Core 模块 → AI Backend → Illustrator.Application COM → Illustrator 引擎
```

### 关键设计原则

| 原则 | 说明 |
|------|------|
| **COM 单例** | 整个会话共享一个 `Illustrator.Application` 实例 |
| **原子操作** | 每个 CLI 命令映射一个或多个 COM API 调用 |
| **JSON 输出** | 所有命令支持 `--json` 标志，输出机器可读格式 |
| **REPL 模式** | 无命令时进入交互式 REPL，支持连续操作 |

## 前置条件

- **Windows 10/11**
- **Adobe Illustrator 2023+**（需注册 COM 接口）
- **Python 3.10+** + `pywin32` + `click`

## 安装

```bash
cd illustrator-harness/agent-harness
pip install -e .
```

## 命令分组

| 命令组 | 描述 | 对应 COM 对象 |
|--------|------|--------------|
| `project` | 新建/打开/保存 AI 项目 | `Application.Documents` |
| `layers` | 图层操作（增删改、可见性、锁定） | `Layer`, `Layers` |
| `shapes` | 形状绘制（矩形、椭圆、线条、多边形） | `PathItems` |
| `text` | 文字操作（添加/修改、字体、大小、颜色） | `TextFrame`, `TextRange` |
| `export` | 导出（PNG、JPEG、SVG、PDF、AI） | `Document.Export` |

## 快速开始

```bash
# 创建新项目
cli-anything-illustrator project new output.ai -w 1920 -h 1080

# 添加文字
cli-anything-illustrator text add "Hello World" --x 100 --y 100 --font "Arial" --size 72

# 绘制矩形
cli-anything-illustrator shapes rect --x 50 --y 50 --w 200 --h 100

# 导出 PNG
cli-anything-illustrator export png output.png

# 获取 JSON 状态（供 Agent 解析）
cli-anything-illustrator --json project info

# 交互式 REPL
cli-anything-illustrator
```

## Claude Code 集成

配合 Claude Code Skill，自然语言即可操控 Illustrator：

```
> 新建 1920x1080 的画板，加标题"品牌设计"，用思源黑体 64pt
> 画一个蓝色矩形 200x150，放在坐标 (100, 100)
> 导出为 SVG 放在桌面上
```

## 许可

MIT License
