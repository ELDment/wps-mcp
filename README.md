<h1 align="center">WPS MCP</h1>

<p align="center">
  <a href="https://pypi.org/project/wps-mcp/"><img src="https://img.shields.io/pypi/v/wps-mcp?style=flat-square&amp;color=2563eb" alt="PyPI 版本"></a>
  <a href="#快速开始"><img src="https://img.shields.io/badge/Python-3.11%2B-3776ab?style=flat-square" alt="Python 3.11 及以上，Linux 需 CPython 3.11"></a>
  <a href="#快速开始"><img src="https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-475569?style=flat-square" alt="支持 Windows、Linux 和 macOS"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-AGPL--3.0--only-7c3aed?style=flat-square" alt="许可证：AGPL-3.0-only"></a>
</p>

<p align="center">
  <strong>🪄 让 AI 接管 WPS，自动处理文档、表格与 PPT</strong>
</p>

<p align="center">
  <strong>45</strong> 个工具组 &nbsp;·&nbsp; <strong>261</strong> 项操作 &nbsp;·&nbsp; 文字 / 表格 / 演示
</p>

<table>
  <tr>
    <td width="50%" valign="top">
      <video src="https://github.com/user-attachments/assets/e8442138-25c2-4f4f-b1a2-e12e2a6628ad" width="100%" controls></video>
    </td>
    <td width="50%" valign="top">
      <video src="https://github.com/user-attachments/assets/b1d0b250-61f4-4a40-aded-55869e5aa8c5" width="100%" controls></video>
    </td>
  </tr>
</table>

---

## 功能与平台对比

**支持三大平台三大办公组件**

| 组件 | 功能 | Windows | Linux | macOS |
| --- | --- | :---: | :---: | :---: |
| **通用** | 环境检测、组件连接与启动 | ✅ | ✅ | ✅ |
| | 文档新建、模板、打开、保存、另存为与关闭 | ✅ | ✅ | ✅ |
| | 文档信息、PDF 导出与页面预览 | ✅ | ✅ | ✅ |
| | 接管用户已打开的 WPS 窗口 | ✅ | ❌ | ✅ |
| **WPS 文字** | 正文读写、查找替换、字体段落与样式管理 | ✅ | ✅ | ✅ |
| | 表格、图片、页面分节与页眉页脚 | ✅ | ✅ | ✅ |
| | 修订跟踪、批注与文档统计 | ✅ | ✅ | ✅ |
| | 书签、目录、脚注尾注与交叉引用 | ✅ | ✅ | ✅ |
| | 本地 CSV 邮件合并 | ✅ | ✅ | ✅ |
| **WPS 表格** | 工作表管理、单元格读写、公式与重算 | ✅ | ✅ | ✅ |
| | 单元格格式、行列调整、合并、冻结与命名区域 | ✅ | ✅ | ✅ |
| | 查找、排序、筛选与结构化表格 | ✅ | ✅ | ✅ |
| | 图表、图表数据源更新与数据透视表 | ✅ | ✅ | ✅ |
| | 条件格式、数据验证与下拉列表 | ✅ | ✅ | 部分⁴ |
| **WPS 演示** | 文稿、模板、属性标签、幻灯片、分节与备注 | ✅ | ✅ | ✅ |
| | 主题、母版、版式、占位符、背景与配色 | ✅ | ✅ | 部分⁴ |
| | 文字与字体排版、Markdown 提取、排版检查与修复 | ✅ | ✅ | ✅ |
| | 形状、图层、样式、连接线、自由曲线与 SmartArt | ✅ | ✅ | 部分⁴ |
| | 图片裁剪、网络图片与图标、音视频与播放设置 | ✅¹ | ✅¹ | 部分⁴ |
| | 表格编辑、图表创建、数据读取与图表格式 | ✅ | ✅ | 部分⁴ |
| | **修改图表数据** | ✅² | ❌ | ❌ |
| | 对象动画与幻灯片切换 | ✅ | ✅ | 部分⁴ |
| | 批注、超链接与窗口视图 | ✅ | ✅ | ✅ |
| | 放映控制 | ✅ | ✅ | ❌⁴ |
| | 撤销、重做与批量操作撤销分组 | ✅ | ✅ | ✅ |
| | PDF、幻灯片图片、形状导出与预览 | ✅ | ✅ | ✅ |
| | 复制图片到剪贴板 | ✅ | ✅³ | ✅ |

- **高级功能的可用性取决于 WPS 版本。**
- **¹ 网络素材与媒体**：网络图片与图标需要联网；媒体裁剪需要当前 WPS 提供对应接口。
- **² 演示图表**：Windows 写入需保持现有分类数、系列数及首工作表标准列布局。Linux 和 macOS 支持创建和读取，暂不支持修改图表数据。含可编辑图表请保存为 `.pptx`。
- **³ Linux 剪贴板**：X11 需 `xclip`，Wayland 需 `wl-copy`。
- **⁴ macOS 部分功能**：自由曲线、SmartArt、连接线和部分返回信息有限制，放映控制暂不支持，详见下方说明。

<details>
<summary><strong>平台兼容性说明</strong></summary>

| 平台 | 功能 | 使用限制 |
| --- | --- | --- |
| **Windows** | 修改演示图表数据 | 保持现有分类数和系列数；数据源需位于首张工作表，第一行为标题、A 列为分类、B 列起为系列数值。 |
| **Linux** | 演示图表 | 支持创建和读取，读取时可能打开图表数据编辑器；暂不支持修改嵌入工作簿。 |
| | 已有窗口 | 使用 MCP 启动的 WPS 实例，不接管已有桌面窗口；请通过工具打开需要编辑的文档。 |
| **macOS** | 演示图表 | 支持创建和读取缓存数值；系列名称返回 `null`，新图表的分类标签可能尚未初始化。暂不支持修改嵌入工作簿。 |
| | 自由曲线与连接线 | 支持直线段和闭合路径，暂不支持贝塞尔曲线及节点读取、编辑。连接线需使用数字连接点索引。 |
| | SmartArt | 支持创建图形和调整布局、颜色、样式；暂不支持节点文字、增删和格式，节点数量返回 `null`。 |
| | 条件格式与主题 | 替换条件格式时保留可读颜色，无法保留不可读取的主题绑定；修改主题配色后，颜色方案名称变为 `WPS MCP`，未指定的颜色保留。 |
| | 部分返回信息 | 母版的标题母版状态、图片的透明色、动画的方向信息返回 `null`，不影响其他已支持的操作。 |
| | 放映控制 | 原生放映接口不可靠，相关操作暂不支持；请使用 WPS 自身的放映控件。 |

</details>

## 快速开始

### 1. 准备环境

安装 **WPS Office** 和 **[uv](https://docs.astral.sh/uv/)**，并确认平台环境满足以下要求：

| 平台 | Python | 原生环境 |
| --- | --- | --- |
| **Windows** | 3.11+ | WPS 已注册 COM 接口 |
| **Linux x86_64** | **CPython 3.11** | WPS Linux、Qt5、已登录的 X11 或带 XWayland 的桌面 |
| **macOS** | 3.11+ | WPS 桌面版，安装包需包含内置自动化驱动 |

### 2. 配置 MCP 客户端

将下文配置添加到 MCP 配置文件中，Windows 和 Linux 推荐通过 **[PyPI](https://pypi.org/project/wps-mcp/)** 接入。macOS 适配尚未发布到 PyPI，请使用下方源码配置。三个平台均使用 `wps-mcp` 命令。

<details open>
<summary><strong>从PyPI接入（Windows / macOS）</strong></summary>

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
<summary>Run in a desktop session with DISPLAY and X authority 报错解决方案</summary>

遇到 `Run in a desktop session with DISPLAY and X authority` 错误时，请在已登录的 Linux 图形桌面会话中启动 MCP 客户端，并将该会话的环境变量传递给 WPS MCP。以下为 Claude Code 和 Codex 的配置示例。

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
<summary><strong>从源码接入</strong></summary>

如需本地开发，克隆项目后按对应平台安装依赖，并将配置中的 `项目的绝对路径` 替换为实际路径。

<details>
<summary><strong>Windows / macOS</strong></summary>

在项目根目录运行：

```sh
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

<details>
<summary>macOS 无法启动或访问文件</summary>

- **找不到 uv**：桌面客户端可能不继承终端的 `PATH`，将 `command` 改为 `command -v uv` 输出的绝对路径。
- **找不到编辑器**：首次连接会通过 WPS 内置模板启动编辑器；若当前版本缺少模板，请先手动打开一个文档。
- **无法读写文件**：WPS 受系统沙箱限制，请先在 WPS 中打开目标文件或选择相应目录，以授予访问权限。
- **存在多个 WPS 进程**：在 MCP 配置的 `env` 中设置 `WPS_DRIVER_SOCKET`，指定 `wps_components` 检测结果中的编辑器端点。若应用容器位置不同，可用 `WPS_DRIVER_DIR` 指定；默认目录为 `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data`。

跨组件操作可能切换编辑器窗口，请避免多个 MCP 服务同时操作同一 WPS 实例。WPS 升级后，建议先用测试文档确认编辑、保存和导出功能。

</details>
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

<details>
<summary>Run in a desktop session with DISPLAY and X authority 报错解决方案</summary>

遇到 `Run in a desktop session with DISPLAY and X authority` 错误时，请在已登录的 Linux 图形桌面会话中启动 MCP 客户端，并将该会话的环境变量传递给 WPS MCP。以下配置使用本地源码，请将 `项目的绝对路径` 替换为实际路径。

**Claude Code**

```json
{
  "mcpServers": {
    "wps": {
      "command": "uv",
      "args": [
        "--directory", "项目的绝对路径",
        "run", "--python", "3.11", "wps-mcp"
      ],
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
command = "uv"
args = [
  "--directory", "项目的绝对路径",
  "run", "--python", "3.11", "wps-mcp"
]
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
<summary><strong>开发检查</strong></summary>

在项目根目录运行：

```sh
uv sync --locked --group dev
uv run ruff check .
uv run ruff format --check .
uv run mypy
uv run pyright
uv build
uv run pytest -m "not native"
```

原生测试需安装 WPS，并设置 `WPS_NATIVE_TESTS=1` 后串行执行 `uv run pytest tests/native`。macOS 需用 `--basetemp` 指定 WPS 可访问的专用测试目录；pytest 可能清空该目录，请勿放入个人文件。

</details>
</details>

---

<p align="right">
  <sub>论坛 · <a href="https://linux.do/t/topic/3064793" title="LINUX DO 社区" target="_blank" rel="noopener noreferrer">LINUX DO</a></sub>
</p>
