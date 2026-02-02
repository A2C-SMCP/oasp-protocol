# OASP - Office AddIn Socket Protocol

> Office AddIn 与后端服务之间的实时通信协议规范

## 概述

OASP (Office AddIn Socket Protocol) 是一个基于 Socket.IO 的通信协议，用于 AI Agent 通过 Office AddIn 控制和操作 Microsoft Office 文档。

## 文档

- **在线文档**: [内部文档服务器地址]
- **本地预览**: `inv docs.serve`

## 快速开始

### 安装依赖

```bash
# 使用 uv 安装
uv sync

# 或使用 pip
pip install -e ".[docs]"
```

### 本地预览文档

```bash
# 启动开发服务器（热重载）
inv docs.serve

# 构建静态文档
inv docs.build
```

## 协议版本

当前版本: **0.1.0**

## 项目结构

```
oasp-protocol/
├── docs/                           # MkDocs 文档源
│   ├── index.md                    # 首页
│   ├── specification/              # 协议规范
│   │   ├── index.md               # 概述
│   │   ├── architecture.md        # 架构设计
│   │   ├── connection.md          # 连接与握手
│   │   ├── events-word.md         # Word 事件定义
│   │   ├── events-ppt.md          # PPT 事件定义 [Draft]
│   │   ├── events-excel.md        # Excel 事件定义 [Draft]
│   │   ├── data-structures.md     # 数据结构
│   │   ├── error-handling.md      # 错误处理
│   │   └── conventions.md         # 通用约定
│   └── appendix/
│       ├── glossary.md            # 术语表
│       └── changelog.md           # 变更日志
├── scripts/                        # 部署脚本
│   └── docs/
├── mkdocs.yml                      # MkDocs 配置
├── pyproject.toml                  # 项目元数据
├── tasks.py                        # Invoke 任务入口
└── README.md
```

## 支持的应用

| 应用 | 状态 | 命名空间 |
|------|------|----------|
| Word | ✅ Stable | `/word` |
| PowerPoint | 📋 Draft | `/ppt` |
| Excel | 📋 Draft | `/excel` |

## 许可证

内部使用

## 维护者

- JQQ <jqq1716@gmail.com>
