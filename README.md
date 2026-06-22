# 客户产品 Excel 汇总工具

按「销售渠道 + 货品名称」对销售单 Excel 进行汇总，将相同货品的「数量」「金额」累计相加，并导出新的 Excel 文件。

## 功能说明

- 读取固定格式的销售单 Excel（如 `销售单查询.xlsx`）
- 保留字段：销售渠道、货品名称、数量、金额
- 合并规则：同一「销售渠道 + 货品名称」合并为一行
- 桌面界面：选择文件、预览汇总结果、导出 Excel

## 环境要求

- Python 3.9+
- Windows（也可在其他支持 tkinter 的系统运行）

界面采用 CustomTkinter 现代化风格，与导出 Excel 表头深绿色主题保持一致。

## 安装依赖

```bash
pip install -r requirements.txt
```

## 运行

```bash
python src/main.py
```

## 输入文件要求

Excel 第一行必须为表头，且包含以下列：

- 销售渠道
- 货品名称
- 数量
- 金额

## 样例数据

项目内提供样例文件：`sample/销售单查询.xlsx`

预期汇总效果：

- 原始行数：29,663
- 汇总后行数：1,478

## 使用步骤

1. 启动程序
2. 点击「选择 Excel」选择销售单文件
3. 点击「开始汇总」
4. 预览结果后点击「导出结果」保存文件

## 打包给同事使用（开箱即用）

同事电脑**无需安装 Python**，只需双击 exe 即可运行。

### 打包步骤（在你自己的电脑上执行一次）

1. 确保已安装 Python 3.10+ 和项目依赖
2. 双击项目根目录下的 [`build.bat`](build.bat)
3. 等待打包完成（约 1–3 分钟）
4. 在 `dist\客户产品Excel汇总工具\` 目录下找到可执行程序

### 发给同事

1. 将 `dist\客户产品Excel汇总工具` **整个文件夹**压缩成 zip
2. 发给同事，解压到任意目录
3. 双击 `客户产品Excel汇总工具.exe` 启动

### 注意事项

- 仅支持 **Windows 64 位** 系统
- 首次启动可能稍慢（约 3–5 秒），属正常现象
- 若杀毒软件误报，添加信任即可（PyInstaller 打包程序偶会被误报）
- 如需自定义图标，可在 `build.spec` 中设置 `icon='app.ico'`

### 命令行打包（可选）

```bash
pip install -r requirements-build.txt
pyinstaller build.spec --noconfirm --clean
```
