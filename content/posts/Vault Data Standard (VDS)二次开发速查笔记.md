+++
title = 'Vault Data Standard (VDS)二次开发速查笔记'
date = 2025-12-16T14:48:40+08:00
draft = false
categories = ['Vault']
keywords = ['Vault','VDS','DataStandard']
+++

## 一、什么是 Autodesk Vault Data Standard（VDS）？

> **常见误解纠正**：  
> 名称 _Data Standard_ 并非指“数据校验（Validation）”，而是 **“数据标准化（Standardization）”** —— 它的核心使命是：  
> **在用户输入源头，将自由、随意、不一致的数据，转化为结构化、规范、可被系统理解的属性值**。

### ▸ 它解决什么问题？

在未启用 Data Standard 时，Vault 用户检入文件仅能填写极少数基础字段（如文件名、生命周期状态），大量业务属性（如 `型号`、`流量`、`材质`、`明细写法`）依赖后续手动补录，极易导致：

- 属性缺失、错填、单位混乱（`1500rpm` vs `1500 RPM`）
- 同类文件规格描述格式不一（`Q=10m³/h, H=20m` vs `10×20`）
- CAD 与 Vault 属性不同步  
  → 最终使得自动化、检索、BOM 生成等高级功能失效。

### ▸ 它如何工作？

当启用 Data Standard 后，在以下关键操作时，**自动弹出一个标准化数据录入窗口**（常称 _Property Dialog_）：

- Vault Client 中：
  - 新建/编辑 **文件**（对应 `File.xaml`）
  - 新建/编辑 **文件夹**（对应 `Folder.xaml`）
- Inventor / AutoCAD 中：
  - **检入（Check In）** 文件时
  - **保存到 Vault** 时

该窗口默认采用 **左右双栏布局**：

| 区域         | 内容                                                                                        | 控制方式                                        |
| ------------ | ------------------------------------------------------------------------------------------- | ----------------------------------------------- |
| **左侧面板** | **固定属性（Static Properties）**<br>由 XAML 手动定义的控件（如 `<TextBox>`、`<ComboBox>`） | 修改 `File.xaml` / `Folder.xaml`                |
| **右侧面板** | **动态属性 Grid（Dynamic Properties）**<br>自动列出该 Vault Category 下关联的所有 Attribute | 由 `DynamicProperties.xml` 控制显隐、顺序、别名 |

> **补充说明**：除文件（File）与文件夹（Folder）外，以下对象也支持 Data Standard（较少用但存在）：
>
> - **Job（任务）**：`Job.xaml`（用于自定义 Job Handler 参数录入）
> - **Change Order（变更单）**：`ChangeOrder.xaml`（需 Vault Professional + Change Order 模块）
> - **BOM View（BOM 视图）**：部分定制场景可用  
>   但 95% 的开发集中在 **File / Folder / CAD 检入** 三类场景。

配合 PowerShell 脚本与窗口生命周期钩子，Data Standard 可实现以下核心能力：

- 自动填充初值（如根据文件名、分类或已有属性推导默认值）
- 动态联动计算（如选中“泵类” → 自动拼接 {流量}×{扬程}×{转速} 生成规格）
- 输入校验与拦截（如必填项检查、单位格式规范、范围值验证）
- UI 动态调整（如按角色/状态控制字段显隐、只读）

## 二、Data Standard 开发的本质：配置覆写而非代码编译

与传统 .NET 插件开发（需编译 DLL、注册 COM、重启服务）不同，**Data Standard 的二次开发本质上是「配置层覆写」**：

- 不新增组件，**只修改/新增特定 XML、XAML、PowerShell 文件**；
- 不依赖构建工具，**改完即部署**（复制 `*.Custom` 目录到目标机器）；
- 不影响原生功能，**升级 Vault 时原配置无损**（Custom 目录独立存在）。

因此，在动手前，**先理解其目录组织逻辑至关重要**——它直接决定了：

- 改哪个文件能影响哪个行为
- 如何打包交付
- 如何确保多环境一致性（开发 / 测试 / 生产环境结构完全对齐）

### ▸ DataStandard 目录结构

| 类型           | 路径                           | 作用                             | 是否可直接修改？             |
| -------------- | ------------------------------ | -------------------------------- | ---------------------------- |
| **原始配置**   | `Vault\`, `CAD\`               | Autodesk 官方默认模板与脚本      | **禁止直接改**（升级易覆盖） |
| **自定义覆盖** | `Vault.Custom\`, `CAD.Custom\` | **二开主战场**：同名文件优先加载 | **推荐只改这里**             |
| **其他文件**   |                                |                                  | **禁止直接改**               |

### ▸ `*.Custom` 内部结构详解（以 `Vault.Custom` 为例）

```text
Vault.Custom\
├── addinVault\                     ← PowerShell 脚本逻辑层（核心！）
│   └── default.ps1             ← 全局钩子脚本（重点）
|   └── ...
├── Configurations\             ← UI 与属性配置层
│   ├── File.xaml               ← 文件属性页模板
│   ├── File.xml
│   ├── Folder.xaml
│   ├── VaultDynamicProperties.xml   ← 动态属性 Grid 控制（关键！）
│   └── ...
└── MenuDefinitions.xml         ← 右键/工具栏菜单定义（独立于 Configurations！）
```

## 三、核心开发点 1：`*DynamicProperties.xml` —— 控制右侧动态属性 Grid

> 这是最轻量的定制入口：仅通过声明式配置，即可统一属性显隐与顺序，为后续所有逻辑打下清晰的数据基础。

### ▸ 作用

- Vault 属性页分两栏：
  - 左边：**静态属性**（硬编码在 XAML 中的 `<TextBox>`）
  - 右边：**动态属性 Grid**（自动列出 Vault 中定义的 Attribute）

### ▸ 控制逻辑

- **未在 `*DynamicProperties.xml` 中定义的 Vault 特性**  
  → 按 Vault 中定义的**原始顺序**显示在**最前面**
- **已定义的 Vault 特性**  
  → 按 `*DynamicProperties.xml` 中 `<Property>` 的**书写顺序**显示在**后面**

### ▸ 示例：屏蔽 `Comment`，调整 `Material` 位置

```xml
<!-- Vault.Custom\Configurations\VaultDynamicProperties.xml -->
<DynamicProperties>
    <DynamicPropertiesCategory Category="工程" Type="FILE">
    <DynamicProperty>设计人</DynamicProperty>
    <DynamicProperty>校对人</DynamicProperty>
    <DynamicProperty Hidden="True">关键词</DynamicProperty>
  </DynamicPropertiesCategory>
</DynamicProperties>
```

## 四、核心开发点 2：`*.xaml` —— 静态属性定制

通过覆写 `CustomVault\Configurations\` 下的 XAML 文件（如 `File.xaml`、`Folder.xaml`、`Job.xaml`），可自定义 Property Dialog 的左侧静态区域。

### ▸ 支持的绑定类型（直接复用 VDS 内建机制）

#### 1. 绑定 Vault Attribute（读写）

```xml
<ComboBox SelectedValue="{Binding Prop[_NumSchm].Value}" ... />
<TextBox Text="{Binding Prop[DETAIL_SPEC].Value}" ... />
```

- `Prop[xxx].Value`：与 Vault 中定义的 Attribute name 严格一致（区分大小写、下划线）
- 支持 `.Text`（显示值）、`.Value`（存储值）分离（用于单位处理）

#### 2. 绑定 PowerShell 数据源

```xml
<ComboBox ItemsSource="{Binding PsList[GetMaterials]}" ... />
<CheckBox IsChecked="{Binding PsValue[IsMFile]}" ... />
```

- `PsList[FunctionName]` → 调用 `default.ps1` 中 `function FunctionName { ... }`，返回数组
- `PsValue[FunctionName]` → 调用同名函数，返回单值（如布尔、字符串）

#### 3. 绑定校验

```xml
<TextBox Text="{WPF:ValidatedBinding Prop[PartNumber].Value}" ... />
```

#### 4. 自定义转换器

除内建转换器外，支持通过自定义 .NET 程序集扩展转换逻辑：

- 编写实现 `IValueConverter` 的类（如 `IncreaseByOneConverter`）
- 编译为 DLL（如 `MyConverters.dll`）
- 将 DLL 置于 **`CustomVault\addins\` 目录**（与 `default.ps1` 同级，VDS 会自动加载该目录下所有 DLL）
- 在 XAML 中声明命名空间并注册为静态资源：

```xml
<WPF:DSWindow xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:my="clr-namespace:MyNamespace;assembly=MyConverters">
>
<Window.Resources>
    <my:IncreaseByOneConverter x:Key="IncreaseByOneConverter" />
</Window.Resources>
```

使用示例：

```xml
<TextBox Text={Binding ..., Converter={StaticResource MyConverter}} />
```

## 五、核心开发点 3：`default.ps1` —— 默认值与生命周期钩子

> 当静态布局不足以表达业务规则时，需引入动态逻辑：通过 PowerShell 钩子，在窗口生命周期关键节点注入计算、回退或校验能力——这是从“能填”迈向“填对”的关键跃升。

这是你提到的“初始化窗口时设初值”的主入口。其核心是 **注册事件处理器**，常用两个阶段：

| 阶段               | 函数                                | 时机                                  | 典型用途                                          |
| ------------------ | ----------------------------------- | ------------------------------------- | ------------------------------------------------- |
| `InitializeWindow` | `function InitializeWindow { ... }` | **窗口打开后、数据绑定后**            | 读取现有值、计算默认值、设置控件状态              |
| `WindowClosing`    | `function WindowClosing { ... }`    | **用户点击“确定/取消”后、窗口关闭前** | 清理资源、记录日志、**阻止关闭**（返回 `$false`） |

### ▸ 示例：自动填充 `明细写法`（DETAIL_SPEC）

```powershell
# CustomVault\addins\default.ps1
function InitializeWindow {
    param($window)

    # 获取属性对象（只读/可写）
    $prop = $window.DataContext.Property

    # 若 DETAIL_SPEC 为空，回退到 DETAIL
    if ([string]::IsNullOrWhiteSpace($prop["DETAIL_SPEC"].Value)) {
        $prop["DETAIL_SPEC"].Value = $prop["DETAIL"].Value
    }

    # 示例：按文件名设 SyncToEPlan（你提过的逻辑）
    if ($window.DataContext.File -and $window.DataContext.File.Name -match "^M") {
        $prop["SyncToEPlan"].Value = "Yes"
    }
}

function WindowClosing {
    param($window, $dialogResult)

    # 若点击“确定”但规格未填写，阻止提交
    if ($dialogResult -eq "OK") {
        $prop = $window.DataContext.Property
        if ([string]::IsNullOrWhiteSpace($prop["DETAIL_SPEC"].Value)) {
            [System.Windows.Forms.MessageBox]::Show("请填写明细规格！", "校验失败", "OK", "Error")
            return $false  # 阻止窗口关闭 → 即阻止检入
        }
    }
    return $true  # 允许关闭
}
```

补充知识点：

- `PreSave` / `PostSave` 是 **JobHandler 或 Vault API** 的钩子，**不是 VDS 的**
- VDS 中只有 `InitializeWindow` / `WindowClosing` 是**窗口级**生命周期
- `$window.DataContext` 包含：`.File`, `.Folder`, `.Property`, `.Vault`, `.User` 等上下文

## 六、核心开发点 4：`MenuDefinitions.xml` —— 给 Vault Client 加按钮

> 前三个层级聚焦**单对象的标准化输入**；而自定义菜单将能力延伸至**多对象批量操作与外部系统集成**，是脱离 Property Dialog、走向自动化流水线的桥梁。

### ▸ 位置与结构

- **路径**：`CustomVault\MenuDefinitions.xml`（不在 Configurations 内！）
- **结构**：

  ```xml
  <MenuDefinitions>
    <MenuItems>
      <MenuItem id="MyExportButton" label="导出规格" ... />
    </MenuItems>

    <CommandSites>
      <!-- 定义按钮出现在哪些上下文 -->
      <CommandSite location="FileContext">        <!-- 文件右键 -->
        <MenuItemRef id="MyExportButton" />
      </CommandSite>
      <CommandSite location="FolderContext">      <!-- 文件夹右键 -->
        <MenuItemRef id="MyExportButton" />
      </CommandSite>
      <CommandSite location="ExplorerToolbar">    <!-- 主工具栏 -->
        <MenuItemRef id="MyExportButton" />
      </CommandSite>
    </CommandSites>
  </MenuDefinitions>
  ```

### ▸ 常用 `location` 对应关系（附图建议位置）

| `location`        | Vault 中位置     | 截图建议标注                |
| ----------------- | ---------------- | --------------------------- |
| `FileContext`     | 文件右键菜单     | 右键任意文件 → 弹出菜单底部 |
| `FolderContext`   | 文件夹右键菜单   | 右键文件夹 → 弹出菜单       |
| `ExplorerToolbar` | 主窗口顶部工具栏 | 左上角“新建”“检入”旁        |
| `PropertyPage`    | 属性页底部按钮区 | 文件属性页 → “确定/取消”旁  |
| `JobQueueContext` | 任务队列右键     | Job Processor 窗口右键      |

> **实现点击逻辑**：  
> 需配合 `addins\default.ps1` 中定义同名函数：
>
> ```powershell
> function MyExportButton_Click {
>     param($selection)  # 当前选中项数组
>     # 你的导出逻辑
> }
> ```
