# FactMesh · 汇报数字核对

在 DSH 中将 PowerPoint 的数字与 Excel 底表逐项核对。点开结果可看到原稿页码、原文、底表单元格、行列标题、期间、单位和比较过程。确认过的关系可在下个月重新定位。

## 安装

验收宿主为 DSH Web `0.2.2-rc.1`，使用带 `workspaceFiles` 的 Web profile。关闭当前 Web 进程后，在 PowerShell 中安装。

```powershell
dsh plugin --profile web add github:jidekaixin2dian/dsh-factmesh
dsh web --host 127.0.0.1 --port 17321
```

用终端打印的地址打开 Web。先添加工作区并打开其中一个会话，然后从左栏进入 **FactMesh · 数字核对**。在全新宿主的引导中，可选择 **添加 API Key** 后点 **稍后配置**。核对无需配置模型。

离线安装时，从 [GitHub Release](https://github.com/jidekaixin2dian/dsh-factmesh/releases) 下载 `.tgz`，将上面的 GitHub 源替换为归档的本地路径。

独立验证可在同一个 PowerShell 中先设置临时 profile 根目录。

```powershell
$env:DSH_HOME = Join-Path $env:TEMP ('factmesh-' + [guid]::NewGuid())
$env:DSH_AGENTS_HOME = Join-Path $env:DSH_HOME 'agents'
```

## 核对一份月报

1. 选择工作区会话和项目名，输入月报期间，例如 `2026-09`。从目录中选择 `.pptx` 和 `.xlsx`。
2. 点 **读取工作表目录**，选好工作表后点 **开始核对**。展开 **解析范围** 可查看扫描页数、对象数、底表单元格数和未覆盖的对象。
3. 在左栏点一条数字，右侧查看候选和底表证据。多个候选时选择具体单元格，也可填地址并点 **查看**。复核业务口径后点 **确认关系**，或暂时跳过。
4. 次期沿用同一浏览器、DSH Web 地址、工作区和项目名，选择下期文件和月份。已复用条目显示首次位置和当前证据。页面标题、业务线、列标题、单位或期间无法唯一对应时，重新确认。

项目关系保存在浏览器本地，按 Web 来源、工作区目录和项目名读取。使用固定端口可在重启后继续复用。浏览器清除站点数据时会一并移除这些关系。Office 文件经宿主只读接口读取。

## 两期样例

解压配套的 `FactMesh-Office-samples.zip`，把四份 Office 文件放到同一个 DSH 工作区。源码仓库内也有 [samples](samples) 目录。两期都选 **经营指标** 工作表。

| 九月待确认的指标 | 应选地址 | 十月结果 |
| --- | --- | --- |
| 电商收入同比 | `F17`，行标题含「电商」 | 从第 6 页移到第 3 页，底表 `D18`。汇报 `12.8%` 比底表 `10.8%` 高 `2.0` 个百分点 |
| 收入环比 | `F19` | `D20`，相符，即使数值与同比相同仍有各自的关系 |
| 营业收入 | `F20` | `D21`，`12,000 万元` 换算为 `1.20 亿元` |
| 毛利率 | `F21` | `D22`，底表 `12.84%` 按汇报一位小数为 `12.8%` |
| 净现金流 | `F22` | `D23`，汇报沿用 `−0.45 亿元`，底表为 `−0.52 亿元` |
| 续费率 | `F23` | 底表改为「续费率（新口径）」，需要手选 `D24` 并确认 |
| 应收比率 | `F24` | `D25`，公式含错误，显示待核实 |

十月 `F17` 已变成「供应链 / 物流费用同比」。两期均有嵌入图表，可在 **解析范围** 中看到未覆盖提示。

## 当前覆盖

- PPTX 支持普通文本框和原生表格。图片、嵌入图表、合并单元格的 PPT 表格及无法可靠识别的对象会列在解析范围。结果提供页码与上下文。
- XLSX 适用于普通值表。将业务线、指标和单位放在数据列左侧，在数据列顶部标明 `YYYY-MM` 及当期、累计、预算等口径。公式、错误值、条件数字格式和无法解释的单位显示待核实。金额支持元、万元和亿元，百分比与百分点分别比较。
- 单文件上限为 24 MiB，单 XML 部件 8 MiB。工作表限制为 5,000 行、128 列、35,000 个单元格，PPTX 限 180 页。遇到上限可拆分文件核对。

## 开发与验收

```powershell
npm ci
npm test
npm run build
New-Item -ItemType Directory -Force dist
npm pack --pack-destination dist
$env:FACTMESH_BROWSER_CHANNEL = 'chrome'
npm run test:host
```

如 CLI 入口不在 Windows npm 的默认路径，设置 `FACTMESH_DSH_CLI` 为已安装 DSH 的 `lib/bin.js`。宿主测试用临时 profile、端口、工作区和样例副本。可用 `FACTMESH_HOST_RESULTS` 指定日志和截图目录。

样例生成器是 `scripts/generate-samples.mjs`。在 Codex 工作区资源中取得 `@oai/artifact-tool` 所在的 `node_modules` 路径，设为 `FACTMESH_ARTIFACT_MODULES`。用该运行时的 Node 执行 `node scripts/generate-samples.mjs work/new-samples` 可生成另一份样例。四份已生成的 Office 文件可直接复跑。

插件身份为 `@jidekaixin2dian/dsh-factmesh`。在 `FACTMESH_MARKET_SOURCE` 指向 dshmarket `1.73.20` 源码与依赖后，`npm run test:update` 通过隔离的 loopback npm 源验证版本检测、官方 DSH 安装和升级后复查。

[English](README.en.md)
