<h1 align="center">WPS MCP</h1>

<p align="center">
  <a href="https://pypi.org/project/wps-mcp/"><img src="https://img.shields.io/pypi/v/wps-mcp?style=flat-square&amp;color=2563eb" alt="PyPI 版本"></a>
  <a href="#快速开始"><img src="https://img.shields.io/badge/Python-3.11%2B-3776ab?style=flat-square" alt="Python 3.11 及以上，Linux 需 CPython 3.11"></a>
  <a href="#快速开始"><img src="https://img.shields.io/badge/Platform-Windows%20%7C%20Linux-475569?style=flat-square" alt="支持 Windows 和 Linux"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-AGPL--3.0--only-7c3aed?style=flat-square" alt="许可证：AGPL-3.0-only"></a>
</p>

<p align="center">
  <strong>🪄 让 AI 接管 WPS，自动处理文档、表格与 PPT</strong>
</p>

<p align="center">
  <strong>45</strong> 个工具组 &nbsp;·&nbsp; <strong>261</strong> 项操作 &nbsp;·&nbsp; 文字 / 表格 / 演示
</p>

<https://github.com/user-attachments/assets/59574cdb-c623-48a9-850d-c4488c5acacc>

---

## 功能与平台对比

**支持两大平台三大办公组件**

| 组件 | 功能 | Windows | Linux |
| --- | --- | :---: | :---: |
| **通用** | 环境检测、组件连接与启动 | ✅ | ✅ |
| | 文档新建、模板、打开、保存、另存为与关闭 | ✅ | ✅ |
| | 文档信息、PDF 导出与页面预览 | ✅ | ✅ |
| | 接管用户已打开的 WPS 窗口 | ✅ | ❌ |
| **WPS 文字** | 正文读写、查找替换、字体段落与样式管理 | ✅ | ✅ |
| | 表格、图片、页面分节与页眉页脚 | ✅ | ✅ |
| | 修订跟踪、批注与文档统计 | ✅ | ✅ |
| | 书签、目录、脚注尾注与交叉引用 | ✅ | ✅ |
| | 本地 CSV 邮件合并 | ✅ | ✅ |
| **WPS 表格** | 工作表管理、单元格读写、公式与重算 | ✅ | ✅ |
| | 单元格格式、行列调整、合并、冻结与命名区域 | ✅ | ✅ |
| | 查找、排序、筛选与结构化表格 | ✅ | ✅ |
| | 图表、图表数据源更新与数据透视表 | ✅ | ✅ |
| | 条件格式、数据验证与下拉列表 | ✅ | ✅ |
| **WPS 演示** | 文稿、模板、属性标签、幻灯片、分节与备注 | ✅ | ✅ |
| | 主题、母版、版式、占位符、背景与配色 | ✅ | ✅ |
| | 文字与字体排版、Markdown 提取、排版检查与修复 | ✅ | ✅ |
| | 形状、图层、样式、连接线、自由曲线与 SmartArt | ✅ | ✅ |
| | 图片裁剪、网络图片与图标、音视频与播放设置 | ✅¹ | ✅¹ |
| | 表格编辑、图表创建、数据读取与图表格式 | ✅ | ✅ |
| | **修改图表数据** | ✅² | ❌ |
| | 对象动画与幻灯片切换 | ✅ | ✅ |
| | 批注、超链接、窗口视图与放映控制 | ✅ | ✅ |
| | 撤销、重做与批量操作撤销分组 | ✅ | ✅ |
| | PDF、幻灯片图片、形状导出与预览 | ✅ | ✅ |
| | 复制图片到剪贴板 | ✅ | ✅³ |

- **高级功能的可用性取决于 WPS 版本。**
- **¹ 网络素材与媒体**：网络图片与图标需要联网；媒体裁剪需要当前 WPS 提供对应接口。
- **² 演示图表**：Windows 写入需保持现有分类数、系列数及首工作表标准列布局。Linux 支持创建和读取，可能打开数据编辑器。含可编辑图表请保存为 `.pptx`。
- **³ Linux 剪贴板**：X11 需 `xclip`，Wayland 需 `wl-copy`。

## 快速开始

### 1. 准备环境

安装 **WPS Office** 和 **[uv](https://docs.astral.sh/uv/)**，并确认平台环境满足以下要求：

| 平台 | Python | 原生环境 |
| --- | --- | --- |
| **Windows** | 3.11+ | WPS 已注册 COM 接口 |
| **Linux x86_64** | **CPython 3.11** | WPS Linux、Qt5、已登录的 X11 或带 XWayland 的桌面 |

### 2. 配置 MCP 客户端

将下文配置添加到 MCP 配置文件中，推荐通过 **[PyPI](https://pypi.org/project/wps-mcp/)** 接入。

<details open>
<summary><strong>从PyPI接入（Windows）</strong></summary>

```json
{
  "mcpServers": {
    "wps": {
      "command": "uvx",
      "args": ["wps-mcp@latest"]
    }
  }
}
```

如需使用预发布版，将 `args` 改为 `["--prerelease", "allow", "wps-mcp@latest"]`。

</details>

<details>
<summary><strong>从PyPI接入（Linux）</strong></summary>

Linux 必须指定 **CPython 3.11**。

```json
{
  "mcpServers": {
    "wps": {
      "command": "uvx",
      "args": ["--no-config", "--python", "3.11", "wps-mcp@latest"]
    }
  }
}
```

如需使用预发布版，将 `args` 改为 `["--no-config", "--python", "3.11", "--prerelease", "allow", "wps-mcp@latest"]`。

<details>
<summary>Run in a desktop session with DISPLAY and X authority</summary>

遇到此错误时，需要将当前 Linux 桌面会话的环境变量转发给 WPS MCP。以下为 Claude Code 和 Codex 的配置示例。

**Claude Code**

```json
{
  "mcpServers": {
    "wps": {
      "command": "uvx",
      "args": ["--no-config", "--python", "3.11", "wps-mcp@latest"],
      "env": {
        "DISPLAY": "${DISPLAY}",
        "XAUTHORITY": "${XAUTHORITY:-}",
        "WAYLAND_DISPLAY": "${WAYLAND_DISPLAY:-}",
        "XDG_RUNTIME_DIR": "${XDG_RUNTIME_DIR:-}",
        "DBUS_SESSION_BUS_ADDRESS": "${DBUS_SESSION_BUS_ADDRESS:-}"
      }
    }
  }
}
```

**Codex**

```toml
[mcp_servers.wps]
command = "uvx"
args = ["--no-config", "--python", "3.11", "wps-mcp@latest"]
env_vars = [
  "DISPLAY",
  "XAUTHORITY",
  "WAYLAND_DISPLAY",
  "XDG_RUNTIME_DIR",
  "DBUS_SESSION_BUS_ADDRESS"
]
```

</details>
</details>

<details>
<summary><strong>从源码接入（可选）</strong></summary>

如需本地开发，克隆项目后按对应平台安装依赖，并将配置中的 `项目的绝对路径` 替换为实际路径。

<details>
<summary><strong>Windows</strong></summary>

在项目根目录运行：

```powershell
uv sync --locked
```

在 MCP 配置文件中添加：

```json
{
  "mcpServers": {
    "wps": {
      "command": "uv",
      "args": [
        "--directory", "项目的绝对路径",
        "run", "wps-mcp"
      ]
    }
  }
}
```

</details>

<details>
<summary><strong>Linux</strong></summary>

在项目根目录运行：

```sh
uv sync --locked --python 3.11
```

在 MCP 配置文件中添加：

```json
{
  "mcpServers": {
    "wps": {
      "command": "uv",
      "args": [
        "--directory", "项目的绝对路径",
        "run", "--python", "3.11", "wps-mcp"
      ]
    }
  }
}
```

</details>
</details>

---

<p align="right">
  <sub>论坛 · <a href="https://linux.do/t/topic/3064793" title="LINUX DO 社区" target="_blank" rel="noopener noreferrer">LINUX DO</a></sub>
</p>
