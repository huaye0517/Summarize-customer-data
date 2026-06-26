# 客户产品 Excel 汇总工具

自动识别两种销售单 Excel 格式，按客户/渠道与货品合并重复行，对「数量」「金额」累计求和，并导出新的 Excel 文件。

## 功能说明

- 自动识别 **内部销售单** 与 **外部销售单** 两种格式
- 桌面界面：选择文件、预览汇总结果、导出 Excel
- 导出表头：灰色底 + 深绿色加粗文字

## 支持的 Excel 格式

程序会根据表头**自动识别**模板，无需手动选择。

### 内部销售单

| 项 | 说明 |
|---|---|
| 识别条件 | 表头含「销售渠道」「货品名称」 |
| 合并规则 | 同一「销售渠道 + 货品名称」合并为一行 |
| 输出列 | 销售渠道、货品名称、数量、金额 |
| 样例 | `sample/销售单查询.xlsx`（29,663 → 1,478 行） |

### 外部销售单

| 项 | 说明 |
|---|---|
| 识别条件 | 表头含「客户」「产品名称」「销售」 |
| 合并规则 | 同一「客户 + 产品名称」合并为一行 |
| 输出列 | 客户、产品名称、数量、金额、销售 |
| 样例 | `sample/5月外部销售数据-惠华.xlsx`（99 → 74 行） |

说明：外部表中「客户」对应内部表的「销售渠道」概念；「销售」列在合并时保留首值。

## 环境要求

- Python 3.9+（需带 tkinter 支持）
- **Windows** 或 **macOS**（Linux 也可通过源码运行，未提供安装包）

界面采用 CustomTkinter 现代化风格，与导出 Excel 表头深绿色主题保持一致。

## 安装依赖

```bash
pip install -r requirements.txt
```

## 运行

```bash
python src/main.py
```

## 使用步骤

1. 启动程序
2. 点击「选择 Excel」选择销售单文件
3. 点击「开始汇总」（程序自动识别表格类型）
4. 预览结果后点击「导出结果」保存文件

## 打包给同事使用（开箱即用）

同事电脑**无需安装 Python**，双击即可运行。

> **重要：** PyInstaller 无法跨平台打包——Windows 的 `.exe` 只能在 Windows 上生成，macOS 的 `.app` 只能在 Mac 上生成。需要在对应系统的电脑上分别打包。

### Windows 打包

1. 确保已安装 Python 3.10+ 和项目依赖
2. 双击项目根目录下的 [`build.bat`](build.bat)
3. 等待打包完成（约 1–3 分钟）
4. 在 `dist\客户产品Excel汇总工具\` 目录下找到可执行程序

发给同事：将 `dist\客户产品Excel汇总工具` **整个文件夹**压缩成 zip，解压后双击 `客户产品Excel汇总工具.exe`。

### macOS 打包（MacBook）

1. 在 Mac 上安装 [Python 3.10+](https://www.python.org/downloads/)（安装包自带 tkinter）
   - 若用 Homebrew：`brew install python-tk@3.12`（需与 Python 版本对应）
2. 在项目根目录执行：

```bash
chmod +x build.sh
./build.sh
```

3. 打包完成后，应用程序位于 `dist/客户产品Excel汇总工具.app`

发给同事：将 `.app` 压缩为 zip 发送；解压后双击打开。若提示「无法验证开发者」，在「系统设置 → 隐私与安全性」中点击「仍要打开」。

### 不打包、直接运行（开发 / 临时使用）

在 Windows 或 Mac 上均可：

```bash
pip install -r requirements.txt
python src/main.py
```

### 注意事项

- Windows 安装包仅适用于 **Windows 64 位**
- macOS 安装包需在 **Apple Silicon 或 Intel Mac** 上分别打包（或在目标架构的 Mac 上构建）
- 首次启动可能稍慢（约 3–5 秒），属正常现象
- Windows 若杀毒软件误报，添加信任即可（PyInstaller 打包程序偶会被误报）
- 如需自定义图标：Windows 在 `build.spec`、macOS 在 `build-mac.spec` 中设置 `icon`

### 命令行打包（可选）

```bash
# Windows
pip install -r requirements-build.txt
pyinstaller build.spec --noconfirm --clean

# macOS
pip install -r requirements-build.txt
pyinstaller build-mac.spec --noconfirm --clean
```
