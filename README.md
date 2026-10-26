<h1 align="center">WPS MCP</h1>

<p align="center">
  🪄 让 AI 接管 WPS Office，自主编辑文档、处理表格和制作演示文稿
</p>

<p align="center">
  <a href="https://pypi.org/project/wps-mcp/"><img src="https://img.shields.io/pypi/v/wps-mcp?style=flat-square&amp;color=2563eb" alt="PyPI 版本"></a>
  <a href="#快速开始"><img src="https://img.shields.io/badge/Python-3.11%2B-3776ab?style=flat-square" alt="Python 3.11 及以上，Linux 需 CPython 3.11"></a>
  <a href="#快速开始"><img src="https://img.shields.io/badge/Windows%20%7C%20Linux%20%7C%20macOS-475569?style=flat-square" alt="Windows、Linux、macOS"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-AGPL--3.0--only-7c3aed?style=flat-square" alt="AGPL-3.0-only"></a>
</p>

<p align="center">
  <strong>45 个工具组 · 252 项操作 · WPS 文字 / 表格 / 演示</strong>
</p>

> WPS MCP 为 AI 助手提供 Word 文档编辑、Excel 表格处理、PowerPoint（PPT）制作及 PDF 导出与预览能力，基于 WPS Office 运行，支持 Windows、Linux 和 macOS。

<table>
  <tr>
    <td width="50%" valign="top">
      <video src="https://github.com/user-attachments/assets/e8442138-25c2-4f4f-b1a2-e12e2a6628ad" width="100%" controls></video>
    </td>
    <td width="50%" valign="top">
      <video src="https://github.com/user-attachments/assets/b1d0b250-61f4-4a40-aded-55869e5aa8c5" width="100%" controls></video>
    </td>
  </tr>
  <tr>
    <td width="50%" valign="top">
      <video src="https://github.com/user-attachments/assets/01be49f9-28d5-4c4c-a78e-df98c388c9cd" width="100%" controls></video>
    </td>
    <td width="50%" valign="top"></td>
  </tr>
</table>

<p align="right"><sub>* 上方视频可能因网络原因无法加载或播放。</sub></p>

## 功能与平台对比

| 组件 | 功能 | Windows | Linux | macOS |
| --- | --- | :---: | :---: | :---: |
| **通用** | 环境检测、组件连接与启动 | ✅ | ✅ | ✅ |
| | 文档新建、模板与打开 | ✅ | ✅ | ✅ |
| | 文档保存、另存为与关闭 | ✅ | ✅ | ✅ |
| | 文档信息读取 | ✅ | ✅ | ✅ |
| | PDF 导出与页面预览 | ✅ | ✅ | ✅ |
| | 接管用户已打开的 WPS 窗口 | ✅ | ❌ | ✅ |
| **WPS 文字** | 正文读写与查找替换 | ✅ | ✅ | ✅ |
| | 字体、段落与样式管理 | ✅ | ✅ | ✅ |
| | 表格与图片 | ✅ | ✅ | ✅ |
| | 页面、分节与页眉页脚 | ✅ | ✅ | ✅ |
| | 修订跟踪、批注与文档统计 | ✅ | ✅ | ✅ |
| | 书签与目录 | ✅ | ✅ | ✅ |
| | 脚注、尾注与交叉引用 | ✅ | ✅ | ✅ |
| | 本地 CSV 邮件合并 | ✅ | ✅ | ✅ |
| **WPS 表格** | 工作表管理 | ✅ | ✅ | ✅ |
| | 单元格读写、公式与重算 | ✅ | ✅ | ✅ |
| | 单元格格式、行列调整、合并与冻结 | ✅ | ✅ | ✅ |
| | 命名区域 | ✅ | ✅ | ✅ |
| | 查找、排序与筛选 | ✅ | ✅ | ✅ |
| | 结构化表格 | ✅ | ✅ | ✅ |
| | 图表与图表数据源更新 | ✅ | ✅ | ✅ |
| | 数据透视表 | ✅ | ✅ | ✅ |
| | 条件格式新增、读取与删除 | ✅ | ✅ | ✅ |
| | 条件格式修改 | ✅ | ✅ | 部分¹ |
| | 数据验证与下拉列表 | ✅ | ✅ | ✅ |
| **WPS 演示** | 文稿、模板与属性标签 | ✅ | ✅ | ✅ |
| | 幻灯片、分节与备注 | ✅ | ✅ | ✅ |
| | 主题应用 | ✅ | ✅ | ✅ |
| | 母版、版式与占位符 | ✅ | ✅ | ✅ |
| | 读取标题母版状态 | ✅ | ✅ | ❌ |
| | 主题配色修改 | ✅ | ✅ | 部分² |
| | 幻灯片背景 | ✅ | ✅ | ✅ |
| | 文字与字体排版 | ✅ | ✅ | ✅ |
| | Markdown 提取、排版检查与修复 | ✅ | ✅ | ✅ |
| | 形状、图层与样式 | ✅ | ✅ | ✅ |
| | 连接线：按数字连接点索引连接 | ✅ | ✅ | ✅ |
| | 连接线：按方向选择连接点 | ✅ | ✅ | ❌ |
| | 自由形状：直线段与闭合路径 | ✅ | ✅ | ✅ |
| | 自由形状：贝塞尔曲线 | ✅ | ✅ | ❌ |
| | 自由形状：节点读取与编辑 | ✅ | ✅ | ❌ |
| | SmartArt 创建、布局、颜色与样式 | ✅ | ✅ | ✅ |
| | SmartArt 节点文字、增删与格式 | ✅ | ✅ | ❌ |
| | 本地图片插入与裁剪 | ✅ | ✅ | ✅ |
| | 读取图片透明色 | ✅ | ✅ | ❌ |
| | 网络图片与图标 | ✅³ | ✅³ | ✅³ |
| | 音视频插入与播放设置 | ✅ | ✅ | ✅ |
| | 音视频裁剪 | ✅⁴ | ✅⁴ | ✅⁴ |
| | 表格编辑 | ✅ | ✅ | ✅ |
| | 图表创建、类型与格式 | ✅⁵ | ✅⁵ | ✅⁵ |
| | 图表数据读取 | ✅ | ✅⁶ | 部分⁶ |
| | **图表数据修改** | 部分⁷ | ❌ | ❌ |
| | 对象动画添加、修改与删除 | ✅ | ✅ | ✅ |
| | 动画信息读取 | ✅ | ✅ | 部分⁸ |
| | 幻灯片切换效果 | ✅ | ✅ | ✅ |
| | 批注、超链接与窗口视图 | ✅ | ✅ | ✅ |
| | 放映控制 | ✅ | ✅ | ❌ |
| | 撤销、重做与批量操作撤销分组 | ✅ | ✅ | ✅ |
| | 幻灯片图片与形状导出 | ✅ | ✅ | ✅ |
| | 复制图片到剪贴板 | ✅ | ✅⁹ | ✅ |

- **实际可用性以当前 WPS 版本为准。**
- **¹** macOS：替换规则时保留可读颜色，无法保留不可读取的主题绑定
- **²** macOS：颜色方案名称会变为 `WPS MCP`，未指定的颜色保留
- **³** 需要联网。
- **⁴** 需要当前 WPS 提供裁剪接口
- **⁵** 含可编辑图表请保存为 `.pptx`
- **⁶** Linux：可能打开图表数据编辑器；macOS：读取缓存数值，系列名称返回 `null`，新图表分类标签可能为空
- **⁷** Windows：需保持现有分类数和系列数，数据源位于首张工作表，第一行为标题、A 列为分类、B 列起为系列数值
- **⁸** macOS：方向信息返回 `null`，其他已支持的信息不受影响
- **⁹** Linux：X11 需 `xclip`，Wayland 需 `wl-copy`

## 快速开始

### 1. 准备环境

先安装 [WPS Office](https://www.wps.cn/) 和 [uv](https://docs.astral.sh/uv/)，再配置 MCP 客户端。

| 平台 | Python | WPS 与桌面环境 | 会话方式 |
| --- | --- | --- | --- |
| Windows | 3.11+ | WPS 已注册 COM 接口 | 可连接已有实例，或按需启动 |
| Linux x86_64 | **CPython 3.11** | WPS Linux、Qt5、已登录的 X11 或带 XWayland 的桌面 | 使用 MCP 管理的实例，不接管已有桌面窗口 |
| macOS | 3.11+ | WPS 桌面版，安装包包含内置自动化驱动 | 连接本机 WPS，首次连接可启动编辑器 |

### 2. 配置 MCP 客户端

#### MCP 配置

Windows / macOS：

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

Linux 必须指定 CPython 3.11：

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

使用预发布版本时，将对应配置中的 `args` 替换为：

```sh
# Windows / macOS
"args": ["--prerelease", "allow", "wps-mcp@latest"]

# Linux
"args": ["--no-config", "--python", "3.11", "--prerelease", "allow", "wps-mcp@latest"]
```

#### 环境配置与故障排查

<details>
<summary>可选环境变量</summary>

通常无需配置。需要指定安装位置、模板或 WPS 实例时，只填写需要的变量，路径使用绝对路径。

| 变量 | 何时设置 |
| --- | --- |
| `WPS_OFFICE_DIR` | Linux 找不到 WPS 时，填写包含 `wps`、`et`、`wpp` 的 `office6` 目录。 |
| `WPS_TEMPLATES_DIR` | 使用自己的模板时，填写模板目录。 |
| `WPS_DRIVER_DIR` | macOS 驱动目录不同于默认的 `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data` 时，填写实际目录。 |
| `WPS_DRIVER_SOCKET` | macOS 需要指定 WPS 实例时，填写检测结果中的 socket 路径，见上方示例。 |

例如，Linux 指定 WPS 安装目录和模板目录（请替换为本机实际路径）：

```json
{
  "mcpServers": {
    "wps": {
      "command": "uvx",
      "args": ["--no-config", "--python", "3.11", "wps-mcp@latest"],
      "env": {
        "WPS_OFFICE_DIR": "/opt/kingsoft/wps-office/office6",
        "WPS_TEMPLATES_DIR": "/home/user/wps-templates"
      }
    }
  }
}
```

已有 `env` 时，将需要的变量合并到同一个 `env` 中。

</details>

<details>
<summary>Linux 桌面环境配置与常见报错</summary>

从已登录的图形桌面启动 MCP 客户端。若出现 `Run in a desktop session with DISPLAY and X authority`，在该桌面的终端查询环境变量：

```sh
env | grep -E '^(DISPLAY|XAUTHORITY|WAYLAND_DISPLAY|XDG_RUNTIME_DIR|DBUS_SESSION_BUS_ADDRESS)='
```

把查询结果填入下面的 `env`。值仅作示例，请替换为实际输出，未输出的项删掉：

```json
{
  "mcpServers": {
    "wps": {
      "command": "uvx",
      "args": ["--no-config", "--python", "3.11", "wps-mcp@latest"],
      "env": {
        "DISPLAY": ":0",
        "XAUTHORITY": "/home/user/.Xauthority",
        "WAYLAND_DISPLAY": "wayland-0",
        "XDG_RUNTIME_DIR": "/run/user/1000",
        "DBUS_SESSION_BUS_ADDRESS": "unix:path=/run/user/1000/bus"
      }
    }
  }
}
```

</details>

<details>
<summary>macOS 启动、权限与多实例</summary>

| 遇到的问题 | 处理方法 |
| --- | --- |
| 找不到编辑器 | 先用 WPS 打开一个文档，再重试连接。 |
| 无法访问文件 | 在 WPS 的“打开”对话框中选择目标文件或目录，授予访问权限。 |
| 检测到多个 WPS 进程 | 让 AI 调用 `wps_components/detect`，从返回的 `editor_sockets` 中选择目标路径，填入 `WPS_DRIVER_SOCKET`。 |

指定 WPS 实例的配置示例，将 socket 路径替换为检测结果中的实际路径：

```json
{
  "mcpServers": {
    "wps": {
      "command": "uvx",
      "args": ["wps-mcp@latest"],
      "env": {
        "WPS_DRIVER_SOCKET": "/Users/user/Library/Containers/com.kingsoft.wpsoffice.mac/Data/wpsdriverlocal_wps_12345"
      }
    }
  }
}
```

同一 WPS 实例请避免同时连接多个 MCP 服务。

</details>

### 本地开发

<details>
<summary>使用本地源码</summary>
下载或克隆项目，在根目录安装依赖：

```sh
# Windows / macOS
uv sync --locked

# Linux
uv sync --locked --python 3.11
```

将下面的 `项目目录` 替换为项目根目录（包含 `pyproject.toml` 的目录），例如 `d:/projects/wps-mcp` 或 `/home/user/projects/wps-mcp`。

Windows / macOS：

```json
{
  "mcpServers": {
    "wps": {
      "command": "uv",
      "args": ["--directory", "项目目录", "run", "wps-mcp"]
    }
  }
}
```

Linux：

```json
{
  "mcpServers": {
    "wps": {
      "command": "uv",
      "args": ["--directory", "项目目录", "run", "--python", "3.11", "wps-mcp"]
    }
  }
}
```

</details>

---

<p align="right">
  <sub>论坛 · <a href="https://linux.do/t/topic/3064793" title="LINUX DO 社区" target="_blank" rel="noopener noreferrer">LINUX DO</a></sub>
</p>
