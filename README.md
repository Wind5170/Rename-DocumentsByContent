# Rename-DocumentsByContent.ps1 开发文档

## 1. 脚本概述

Rename-DocumentsByContent.ps1 是一个功能强大的 PowerShell 脚本，用于自动识别多种文档格式中的发文号和发文名，并根据识别结果重命名文件。

### 主要功能：

- 支持多种文档格式：.docx, .doc, .wps, .pdf, .ofd, .gd, .ceb, .cebx
- 集成 OCR 技术处理扫描文档
- 智能识别发文号和发文名
- 灵活的文件类型筛选和跳过机制
- 详细的日志和错误处理
- 支持批量处理和递归处理

### 应用场景：

- 公文管理系统的文件重命名
- 文档归档和整理
- 批量处理大量办公文档

## 2. 系统架构

### 2.1 模块结构

```
├── 主脚本 (Rename-DocumentsByContent.ps1)
├── 模块文件
│   ├── Config.ps1 - 配置管理模块
│   ├── Cache.ps1 - 缓存管理模块
│   ├── Export.ps1 - 导出功能模块
│   └── Test-ExtractDocTitle.ps1 - 测试模块
├── 配置文件
│   ├── config.json - 主配置文件
│   └── custom-rules.json - 自定义规则文件
├── 文档处理引擎模块
│   ├── DOCX处理子模块
│   ├── PDF处理子模块  
│   ├── DOC/WPS处理子模块
│   ├── OFD处理子模块
│   ├── GD处理子模块
│   └── CEB处理子模块
├── OCR识别模块
├── 发文号提取与清洗模块
├── 发文名提取模块
└── 文件重命名模块
```

### 2.2 核心流程

1. **参数解析**：解析命令行参数，设置默认值
2. **环境检查**：检查脚本目录和依赖项
3. **文件扫描**：获取文件列表，应用筛选条件
4. **核心处理**：
   - 文档内容提取
   - 发文号识别与清洗
   - 发文名提取
   - 文件名比较
   - 文件重命名
5. **结果统计**：显示处理结果
6. **资源清理**：释放资源，清理临时文件

## 3. 核心功能模块

### 3.1 文档处理引擎

**Get-DocumentText**：主分发函数，根据文件类型调用相应处理函数

| 格式         | 处理函数           | 依赖                         |
| ---------- | -------------- | -------------------------- |
| .docx      | Get-DocxText   | Xceed.Words.NET 或 DocX.dll |
| .pdf       | Get-PdfText    | iTextSharp + Tesseract OCR |
| .doc/.wps  | Get-DocComText | Word/WPS COM 接口            |
| .ofd       | Get-OfdText    | 解压 + XML提取 + OCR           |
| .gd        | Get-GdText     | Sursen Reader 自动化          |
| .ceb/.cebx | Get-CebText    | Apabi Reader 自动化           |

### 3.2 OCR 识别模块

**智能 OCR 策略**：

1. **基础识别**：直接使用 Tesseract 进行快速识别
2. **优化识别**：使用 ImageMagick 进行图像预处理后再次识别

**关键函数**：

- Invoke-TesseractBase：基础 OCR 识别
- Invoke-TesseractOptimized：带图像预处理的 OCR
- Invoke-SmartOCR：智能调度两阶段 OCR

### 3.3 发文号处理模块

**处理流程**：

1. **直接模式匹配**：使用正则表达式直接匹配全文
2. **逐行提取**：逐行处理，利用历史行组合提高识别率
3. **标准化**：统一括号格式，处理空白字符
4. **清洗**：去除文件头，确保标准格式

**关键函数**：

- NormalizeDocNumber：发文号初步标准化
- CleanDocNumber：发文号最终清洗
- ExtractDocNumberFromLine：从文本行提取发文号

### 3.4 发文名提取模块

**处理流程**：

1. **定位发文号**：在文本中找到发文号位置
2. **多行重组**：组合多行内容，尝试识别发文名
3. **模式匹配**：按优先级尝试多种正则表达式模式
4. **完整性检查**：确保发文名符合完整性要求

**关键函数**：

- ExtractDocTitle：提取发文名（核心函数）
- Test-DocTitleIntegrity：检查发文名完整性

### 3.6 文件重命名模块

**处理逻辑**：

1. **文件名生成**：组合发文号和发文名
2. **冲突处理**：处理文件名冲突
3. **执行重命名**：调用 Rename-Item 执行文件重命名

**关键函数**：

- Format-FileName：格式化文件名中的非法字符
- CompareDocTitleWithFilename：比较发文名与原文件名

### 3.7 配置文件系统

**配置文件**：

- `config.json`：主配置文件，包含工具路径、处理参数等
- `custom-rules.json`：自定义规则文件，包含自定义的发文号和发文名模式

**配置文件结构**：

**config.json**：

```json
{
  "TesseractPath": "C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
  "GhostscriptPath": "C:\\Program Files\\gs\\gs10.06.0\\bin\\gswin64c.exe",
  "MagickPath": "C:\\Program Files\\ImageMagick-7.1.2-15-portable-Q16-x64\\magick.exe",
  "MaxLines": 10,
  "PdfPagesToRead": 1,
  "MaxNestedLevel": 3,
  "DocTypes": ["通知", "报告", "请示", "批复", "函", "意见", "决定", "命令", "指示", "通报", "决议", "公告", "通告"],
  "SkipPrefixes": ["附件", "附表"],
  "SkipKeywords": ["汇编", "选编", "模板", "简报"],
  "GdProcessMode": "Default",
  "SendKeysMethod": "API",
  "CustomDocNumberPatterns": [],
  "CustomDocTitlePatterns": []
}
```

**custom-rules.json**：

```json
{
  "CustomDocNumberPatterns": [
    "[A-Za-z0-9]+\\s*[\\(（\\[\\[]\\s*\\d{4}\\s*[\\)）\\]\\]]\\s*\\d+\\s*号",
    "[\\u4e00-\\u9fa5]+\\s*[令命]\\s*第\\s*\\d+\\s*号"
  ],
  "CustomDocTitlePatterns": [
    "关于\\s*[制定|发布|实施|印发|转发|批转]\\s*.*\\s*的\\s*[通知|报告|请示|批复|函|意见|决定|命令|指示|通报]",
    "[A-Za-z0-9]+\\s*关于\\s*.*\\s*的\\s*[通知|报告|请示|批复|函|意见|决定|命令|指示|通报]"
  ],
  "CustomSkipPrefixes": ["草稿", "讨论稿", "征求意见稿"],
  "CustomSkipKeywords": ["内部使用", "请勿外传", "保密"]
}
```

## 4. 关键函数详解

### 4.1 ExtractDocTitle 函数

**功能**：从文本中提取发文名，支持多层嵌套和多行重组

**参数**：

- `$Text`：文档文本内容
- `$DocNumber`：提取到的发文号
- `$MaxNestedLevel`：最大嵌套层数，默认值为 3

**核心算法**：

1. **发文号定位**：从发文号之后开始提取内容
2. **文本清理**：去除空白字符和特殊字符
3. **多行重组**：组合 1-10 行文本，构建完整的发文名
4. **模式匹配**：按优先级尝试多种正则表达式模式
5. **完整性检查**：确保发文名符合完整性要求
6. **智能处理**：处理书名号配对，确保提取完整

**关键特性**：

- 保留机构名称
- 灵活处理带空格的发文号
- 支持多种转发文件格式
- 处理并列书名号和嵌套书名号
- **动态嵌套层数控制**：根据 `MaxNestedLevel` 参数动态调整正则表达式模式

#### MaxNestedLevel 参数与嵌套正则的协调处理

**设计原理**：
`MaxNestedLevel` 参数用于控制脚本能够识别的最大嵌套层数，与嵌套正则表达式协同工作，实现对复杂公文格式的精确匹配。

**嵌套层级定义**：

- **单层嵌套（Level 1）**：`关于印发《XX》的通知`
- **双层嵌套（Level 2）**：`关于转发《XX关于印发〈YY〉的通知》的通知`
- **三层嵌套（Level 3）**：`关于转发《XX关于印发〈YY关于印发〈ZZ〉的通知〉的通知》的通知`

**动态模式构建机制**：

```powershell
# 基础模式（始终包含）
$basePatterns = @(
    "[\u4e00-\u9fa5a-zA-Z0-9，、]*关于(?:印发|转发|批转).*?的(?:$docTypesPattern)",
    "关于(?:印发|转发|批转).*?的(?:$docTypesPattern)",
    # ... 其他基础模式
)

# 根据 MaxNestedLevel 动态添加嵌套模式
$patterns = $basePatterns.Clone()

# 添加双层嵌套模式（如果 MaxNestedLevel >= 2）
if ($MaxNestedLevel -ge 2) {
    $patterns += "[\u4e00-\u9fa5a-zA-Z0-9，、]*关于(?:印发|转发|批转)《[\u4e00-\u9fa50-9()（）]*〈[\u4e00-\u9fa50-9()（）]+〉[\u4e00-\u9fa50-9()（）]*》的(?:$docTypesPattern)"
}

# 添加三层嵌套模式（如果 MaxNestedLevel >= 3）
if ($MaxNestedLevel -ge 3) {
    $patterns += "[\u4e00-\u9fa5a-zA-Z0-9，、]*关于(?:印发|转发|批转)《[\u4e00-\u9fa50-9()（）]*〈[\u4e00-\u9fa50-9()（）]*〈[\u4e00-\u9fa50-9()（）]+〉[\u4e00-\u9fa50-9()（）]*〉[\u4e00-\u9fa50-9()（）]*》的(?:$docTypesPattern)"
}
```

**使用示例**：

```powershell
# 只识别单层嵌套（提高性能，减少误匹配）
.\Rename-DocumentsByContent.ps1 -FolderPath "D:\文档" -MaxNestedLevel 1

# 识别单层和双层嵌套（平衡性能和准确性）
.\Rename-DocumentsByContent.ps1 -FolderPath "D:\文档" -MaxNestedLevel 2

# 识别最多三层嵌套（处理复杂公文格式，默认设置）
.\Rename-DocumentsByContent.ps1 -FolderPath "D:\文档" -MaxNestedLevel 3
```

**性能与准确性权衡**：

- **MaxNestedLevel = 1**：匹配速度最快，适合处理简单格式的公文
- **MaxNestedLevel = 2**：平衡性能和准确性，适合大多数场景
- **MaxNestedLevel = 3**：能够处理最复杂的嵌套格式，但可能增加误匹配风险

**注意事项**：

- 当前脚本最多支持 3 层嵌套，设置大于 3 的值将被限制为 3
- 增加嵌套层数会增加正则表达式的复杂度，可能略微影响处理速度
- 建议根据实际文档的复杂度选择合适的嵌套层数

### 4.2 Test-DocTitleIntegrity 函数

**功能**：检查发文名的完整性

**参数**：

- `$Title`：要检查的发文名
- `$DocTypes`：文种列表

**检查规则**：

1. **书名号配对**：确保书名号成对出现
2. **结束条件检查**：
   - 情况 1：以"的+文种"结尾（如"的通知"）
   - 情况 2：以文种结尾（用于转发文件格式）
   - 情况 3：以书名号结尾（用于转发文件格式）

### 4.3 Format-FileName 函数

**功能**：格式化文件名中的非法字符

**参数**：

- `$FileName`：原始文件名

**处理步骤**：

1. **转换半角单书名号**：将 `<` 和 `>` 转换为 `＜` 和 `＞`
2. **去除无效字符**：去除文件系统不允许的字符
3. **清理空格**：去除多余空格，确保文件名整洁
4. **处理边界**：去除开头和结尾的分隔符
5. **空值处理**：确保文件名不为空

### 4.4 Get-StringSimilarity 函数

**功能**：计算两个字符串的相似度，处理 OCR 识别错误

**参数**：

- `$String1`：第一个字符串
- `$String2`：第二个字符串

**算法**：使用 Levenshtein 距离算法计算相似度，返回值范围为 0-1，1 表示完全相同

## 5. 技术要点

### 5.1 智能 OCR 策略

**两阶段识别**：

1. **基础识别**：快速识别，适合清晰文档
2. **优化识别**：图像预处理 + OCR，适合模糊文档

**图像预处理步骤**：

- 灰度转换
- 对比度增强
- 自适应锐化
- 噪点去除

### 5.2 发文号提取增强

**多层次匹配策略**：

1. **单行匹配**：对每行文本应用多种正则模式
2. **历史行组合**：结合最近的历史行进行跨行匹配
3. **智能清洗**：包含文件头识别和去除逻辑

**文件头识别模式**：

- 精确匹配模式（如"中国人民银行文件"）
- 通用机构模式（如"XX文件"、"XX办公厅"）
- 发文代字识别（如"银办发"、"南银发"等）

### 5.3 发文名识别优化

**关键技术**：

- **多行重组**：组合多行文本，提高识别率
- **模式优先级**：按优先级尝试多种正则模式
- **书名号处理**：支持并列书名号和嵌套书名号
- **机构名称保留**：保留发文名前的机构名称
- **灵活的发文号匹配**：处理带空格的发文号
- **动态嵌套层数控制**：根据 `MaxNestedLevel` 参数动态调整匹配深度

#### 嵌套正则表达式模式详解

脚本支持三种主要的嵌套模式，通过 `MaxNestedLevel` 参数控制：

**1. 并列引用模式（始终启用）**

```
格式：关于印发《A》和《B》的通知
正则：[\u4e00-\u9fa5a-zA-Z0-9，、]*关于(?:印发|转发)《[\u4e00-\u9fa50-9()（）\[\]]+》[和、]?《[\u4e00-\u9fa50-9()（）\[\]]+》的(?:$docTypesPattern)
```

**2. 双层嵌套模式（MaxNestedLevel >= 2）**

```
格式：关于转发《XX关于印发〈YY〉的通知》的通知
正则：[\u4e00-\u9fa5a-zA-Z0-9，、]*关于(?:印发|转发|批转)《[\u4e00-\u9fa50-9()（）]*〈[\u4e00-\u9fa50-9()（）]+〉[\u4e00-\u9fa50-9()（）]*》的(?:$docTypesPattern)
```

**3. 三层嵌套模式（MaxNestedLevel >= 3）**

```
格式：关于转发《XX关于印发〈YY关于印发〈ZZ〉的通知〉的通知》的通知
正则：[\u4e00-\u9fa5a-zA-Z0-9，、]*关于(?:印发|转发|批转)《[\u4e00-\u9fa50-9()（）]*〈[\u4e00-\u9fa50-9()（）]*〈[\u4e00-\u9fa50-9()（）]+〉[\u4e00-\u9fa50-9()（）]*〉[\u4e00-\u9fa50-9()（）]*》的(?:$docTypesPattern)
```

#### 模式匹配优先级

脚本按照以下优先级尝试匹配模式：

1. **\[机构]关于(印发|转发|批转)...的文种** - 保留机构名称的完整格式
2. **关于(印发|转发|批转)...的文种** - 标准格式
3. **\[机构]关于...的文种** - 简化格式（保留机构）
4. **关于...的文种** - 简化格式
5. **(印发|转发|批转)...的文种** - 极简格式
6. **\[机构]转发\[书名号]...** - 转发文件格式（保留机构）
7. **转发\[书名号]...** - 转发文件格式
8. **\[机构]关于(印发|转发)《A》和《B》...的文种** - 并列引用格式
9. **\[机构]关于(印发|转发)《...〈...〉...》...的文种** - 双层嵌套格式（Level 2+）
10. **\[机构]关于(印发|转发)《...〈...〈...〉...〉...》...的文种** - 三层嵌套格式（Level 3+）

#### 书名号配对检查

在匹配过程中，脚本会动态检查书名号配对情况：

```powershell
# 检查书名号配对
$leftCount = ($matchedTitle.ToCharArray() | Where-Object { $_ -eq '《' }).Count
$rightCount = ($matchedTitle.ToCharArray() | Where-Object { $_ -eq '》' }).Count
$leftSingleCount = ($matchedTitle.ToCharArray() | Where-Object { $_ -eq '〈' }).Count
$rightSingleCount = ($matchedTitle.ToCharArray() | Where-Object { $_ -eq '〉' }).Count

# 如果书名号不配对，继续向后查找下一个文种
if (($leftCount -ne $rightCount) -or ($leftSingleCount -ne $rightSingleCount)) {
    # 尝试修复或继续查找
}
```

### 5.4 错误处理与容错

**多层次错误处理**：

- 函数级 try-catch 块
- 优雅的降级策略
- 详细的错误信息和建议

**资源管理**：

- 临时文件和目录的自动清理
- COM 对象的正确释放
- 内存使用优化

## 6. 部署与使用

### 6.1 系统要求

- PowerShell 5.1 或更高版本
- .NET Framework 4.5 或更高版本

### 6.2 依赖项

**可选依赖**：

| 功能        | 依赖软件                 | 说明                 |
| --------- | -------------------- | ------------------ |
| Word 文档处理 | Microsoft Word 2010+ | 处理 .doc 文件         |
| WPS 文档处理  | WPS Office           | 处理 .wps 文件         |
| OFD 文档处理  | OFD 阅读器              | 处理 .ofd 文件         |
| GD 文档处理   | Sursen Reader        | 处理 .gd 文件          |
| CEB 文档处理  | 方正 Apabi Reader      | 处理 .ceb 和 .cebx 文件 |
| OCR 识别    | Tesseract OCR        | 用于 PDF 文本提取        |

**DLL 依赖**：

- iTextSharp.dll - PDF 处理
- Xceed.Words.NET.dll - Word 文档处理
- BouncyCastle.Cryptography.dll - 加密支持
- System.IO.Packaging.dll - 打包支持

### 6.3 基本用法

**参数说明**：

- `FolderPath`：要处理的文件夹路径
- `Recurse`：是否递归处理子文件夹
- `FileType`：文件类型筛选
- `SkipPrefixes`：跳过以指定前缀开头的文件
- `SkipKeywords`：跳过包含指定关键词的文件
- `DocTypes`：文种列表，用于识别发文名
- `MaxNestedLevel`：最大嵌套层数
- `EnableDebug`：是否启用调试模式
- `PreviewOnly`：是否仅预览，不执行实际重命名
- `EnableCache`：是否启用缓存
- `RunTests`：是否运行测试
- `ExportToCsv`：是否导出结果到 CSV
- `ExportToExcel`：是否导出结果到 Excel

**示例**：

```powershell
# 处理所有文件，跳过附件
.Rename-DocumentsByContent.ps1 -FolderPath "D:\文档" -SkipPrefixes "附件", "附表"

# 仅处理 PDF 文件
.Rename-DocumentsByContent.ps1 -FolderPath "D:\文档" -FileType pdf

# 处理多层嵌套格式
.Rename-DocumentsByContent.ps1 -FolderPath "D:\文档" -MaxNestedLevel 3

# 仅预览，不执行重命名
.Rename-DocumentsByContent.ps1 -FolderPath "D:\文档" -PreviewOnly

# 运行测试
.Rename-DocumentsByContent.ps1 -RunTests

# 导出结果到 CSV
.Rename-DocumentsByContent.ps1 -FolderPath "D:\文档" -ExportToCsv
```

## 7. 测试与维护

### 7.1 测试模式

脚本支持测试模式，通过 `-RunTests` 参数启用。测试模式会执行预设的测试用例，验证核心功能的正确性。

**测试模块**：

- `Test-ExtractDocTitle.ps1`：专门的测试模块，包含各种测试用例

**测试用例**：

- 包含当前文档发文号的多行文本
- 正常转发文件格式
- 带有空格的转发文件格式
- 多行转发文件格式
- 包含转发文件发文号的格式
- 包含半角单书名号的转发文件
- 只有半角单书名号
- 混合使用半角和全角书名号
- 空文本情况
- 只有空格的文本

**测试命令**：

```powershell
# 运行所有测试
.Rename-DocumentsByContent.ps1 -RunTests
```

### 7.2 日志系统

脚本会生成详细的日志文件，记录处理过程和结果：

- 日志文件位置：`Logs\RenameLog_yyyyMMdd_HHmmss.txt`
- 日志内容：处理状态、错误信息、重命名记录

### 7.3 常见问题与解决方案

| 问题                   | 可能原因                   | 解决方案                      |
| -------------------- | ---------------------- | ------------------------- |
| COM 接口失败             | Office/WPS 未正确安装       | 尝试以管理员身份运行                |
| OCR 识别率低             | Tesseract 未安装或配置错误     | 检查 Tesseract 路径和语言包       |
| 文件访问拒绝               | 权限不足或文件被锁定             | 关闭占用程序或提升权限               |
| CEB 处理失败             | Apabi Reader 未安装或版本不兼容 | 安装 Apabi Reader 4.0+      |
| GD 处理失败              | Sursen Reader 未安装      | 安装 Sursen Reader          |
| 处理完文档后需要按 Enter 键    | 脚本执行方式问题               | 已修复，使用 Start-Process 后台执行 |
| Resolve-Path 找不到路径   | 文件已被重命名，仍尝试访问原路径       | 已修复，更新缓存逻辑使用新路径           |
| 无法绑定参数 'Name' 因为它是空值 | 文件名生成失败                | 已修复，添加文件名验证检查             |
| 路径格式不支持              | 脚本参数传递错误               | 已修复，使用数组格式传递参数            |

## 8. 代码优化建议

### 8.1 性能优化

1. **并行处理**：对于大量文件，考虑使用 PowerShell 的并行处理能力
2. **缓存机制**：已实现，使用 cache.json 缓存处理结果
3. **延迟加载**：按需加载依赖项，减少启动时间

### 8.2 功能增强

1. **配置文件**：已实现，支持从 config.json 和 custom-rules.json 读取配置
2. **批量预览**：已实现，通过 -PreviewOnly 参数支持
3. **自定义规则**：已实现，通过 custom-rules.json 支持
4. **导出功能**：已实现，支持导出处理结果到 CSV 或 Excel

### 8.3 代码质量

1. **模块化**：已实现，拆分为多个模块文件
2. **注释**：完善函数文档和关键代码注释
3. **单元测试**：已实现，通过 Test-ExtractDocTitle.ps1 模块
4. **错误处理**：增强错误处理，提高脚本健壮性

## 9. 总结

Rename-DocumentsByContent.ps1 是一个功能强大、设计完善的文档重命名工具，通过智能识别和处理技术，能够高效地处理各种格式的文档文件。

**核心优势**：

- 多格式支持和引擎兼容（支持 .docx, .doc, .wps, .pdf, .ofd, .gd, .ceb, .cebx）
- 智能 OCR 和发文号提取算法
- 完善的错误处理和用户友好界面
- 灵活的配置选项和跳过机制
- 强大的发文名识别能力（保留机构名称、处理带空格的发文号、支持多种转发格式）
- 模块化设计，代码结构清晰
- 配置文件系统，支持自定义规则
- 缓存机制，提高重复处理效率
- 导出功能，支持 CSV 和 Excel 格式
- 测试模块，确保功能稳定性
- 遵循 PowerShell 最佳实践

**已修复的问题**：

- 处理完文档后需要按 Enter 键的问题
- Resolve-Path 找不到路径的问题
- 无法绑定参数 'Name' 因为它是空值的问题
- 路径格式不支持的问题

通过遵循本文档中的指导，开发人员可以更好地理解、维护和扩展这个工具，以满足不同场景的文档处理需求。
