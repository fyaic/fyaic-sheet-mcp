# fyaic-sheet-skill

**PowerShell脚本 + 自然语言指令** = 简单高效的工作表解析方案

### 表格理解增强
- ✅ 自动识别表头层级
- ✅ 理解合并单元格语义
- ✅ 还原树状结构（大阶段 → 子板块 → 具体动作）
- ✅ 处理空单元格和超链接

## Skill架构

```mermaid
flowchart TD
    subgraph Index串联
        A[PowerShell脚本] --> B[Excel文件解析能力]
        C[自然语言指令] --> D[学习复杂表格排布阅读方法]
        B --> E[结构化数据提取]
        D --> F[排版含义分析]
    end
    
    subgraph 功能
        E --> I[梳理多页Sheet目录]
        I --> J{工作簿是否庞大?}
        J -- 是 --> K[反问用户缩小任务范围]
        J -- 否 --> L[直接确定需要阅读的Sheet]
        K --> M[确定具体工作表]
        L --> M
        M --> G[解析工作表]
        F --> H[理解排版逻辑]
        G --> H
    end

```


## 快速开始

### 1. 准备环境
- **Windows**: PowerShell已内置
- **macOS/Linux**: 安装PowerShell Core
  ```bash
  brew install powershell  # macOS
  sudo apt install powershell  # Ubuntu/Debian
  ```

### 2. 克隆项目
```bash
git clone https://github.com/fyaic/fyaic-sheet-skill.git
cd fyaic-sheet-skill
```


## 工具使用流程

推荐入口脚本为 SkillMain.ps1。调用时优先运行该脚本，脚本会先输出 prompts/通用指令.md 的全部内容，再根据工作簿的工作表数量引导你选择要读取的工作表，并最终调用内部的脚本完成数据读取。

```mermaid
flowchart TD
    A[用户/模型] --> B[运行 SkillMain.ps1]
    B --> C[输出通用指令]
    C --> D[调用 ListSheets 获取工作表清单]
    D --> E{表格页数是否庞大?}
    E -- 是 --> F[引导用户缩小任务范围]
    E -- 否 --> G[向用户确认需要阅读的sheet]
    F --> H[确定具体工作表]
    G --> H
    H --> I[调用 ReadSheets 读取数据]
    I --> J[获取并分析数据]
    J --> K[返回结果给用户]
```

