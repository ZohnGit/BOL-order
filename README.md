# BOL 订单导出工具

这是一个 Windows 桌面应用，用于按照提供的 n8n 工作流逻辑读取 bol.com Retailer API 订单，并输出取消单信息和正常订单 XLSX。

## 功能

- 使用 bol.com API User/Password 登录。
- User 和 Password 按需求在界面中明文显示。
- 登录成功后自动将 User/Password 组合保存到 Windows 凭据管理器。
- 支持选择已保存的组合，也可以整对删除。
- 登录成功后显示“选择日期”。
- 按原工作流分页读取 `FBB`、`ALL` 状态订单。
- 原工作流 `If` False 出口的整单取消订单显示在结果表中。
- 正常订单逐单获取订单详情、买家资料及物流资料。
- 显示运行进度；只有 XLSX 在内存中成功生成后才显示 100%。
- 无需安装 Excel 即可生成 `.xlsx` 文件。

## 与原工作流一致的关键规则

- 从第 1 页开始查询。
- 最后一条订单日期大于或等于所选日期时继续翻页。
- 商品净数量为 `quantity - quantityCancelled`。
- 所有商品行净数量都不大于 0 时，订单进入 `If` False 取消单出口。
- 取消单结果位于原工作流的日期筛选之前，因此保持原出口范围，不额外按所选日期过滤。
- 非整单取消订单累计后，只保留下单日期等于所选日期的订单。
- 每单处理前等待 1.5 秒。
- 第 1、16、31……单重新登录获取 Token。
- Shipment 列表使用第一条 `shipmentId` 获取详情。
- 如果订单没有 `shipmentId`，保留该订单并继续处理下一单；XLSX 中物流跟踪号留空。
- `shipmentDateTime` 沿用原工作流，来源为商品的 `latestChangedDateTime`。

## 直接在 GitHub 打包

1. 新建一个 GitHub 仓库。
2. 将本目录中的全部文件上传到仓库根目录，包括 `.github` 文件夹。
3. 打开仓库的 **Actions** 页面。
4. 选择 **Build Windows EXE**。
5. 点击 **Run workflow**。
6. 构建完成后，在该次运行页面的 **Artifacts** 下载 `BolOrderExporter-win-x64`。
7. 解压后得到 `BolOrderExporter.exe`。

GitHub Actions 生成的是 Windows x64 自包含单文件，目标电脑不需要预先安装 .NET 运行时。

## 本地构建（Windows）

安装 [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) 后，在项目目录执行：

```powershell
dotnet publish .\BolOrderExporter.csproj `
  --configuration Release `
  --runtime win-x64 `
  --self-contained true `
  --output .\publish\win-x64 `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:EnableCompressionInSingleFile=true
```

生成文件位于：

```text
publish\win-x64\BolOrderExporter.exe
```

## 使用方法

1. 打开应用，输入或选择 User 和 Password。
2. 点击“登录”。只有 bol.com API 登录成功才会保存凭证并进入下一步。
3. 在“选择日期”页面设置要处理的日期。
4. 点击“开始运行”。
5. 等待进度达到 100%。
6. 查看或复制取消单结果，点击“获取 XLSX 文件”保存正常订单表格。

## 数据与安全

- Password 按明确需求在应用输入框中明文显示。
- 保存的账号组合使用 Windows 凭据管理器，不写入项目文件或普通配置文件。
- 删除账号会删除对应的 User 和 Password 本地凭据。
- API 凭据、Token 和买家资料不会写入应用日志。
- XLSX 含有买家姓名、邮箱等个人信息，请妥善保管。

## API 与错误处理

应用使用工作流中的接口和 `application/vnd.retailer.v10+json` 媒体类型。如果登录、订单接口、Shipment 接口或 XLSX 生成失败，进度不会显示 100%，并会显示失败阶段；失败结果不会开放 XLSX 保存按钮。
