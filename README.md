# mysheet MCP 服务器

Mysheet-mcp 是一个用于将 Excel 文件转换为 JSON 格式的 MCP 服务器，特别适用于在大型模型或智能体中进行数据处理。本项目支持解析 Excel 文件，并将 Excel 中嵌入的文件自动上传到腾讯云 COS（对象存储）。

## 主要功能

- **Excel 转 JSON**：将上传的 Excel 文件转换为 JSON 格式，包含详细的单元格类型信息（文本、数字、货币、日期等）。
- **行对象模式**：支持 "row-object" 解析模式，将每一行视为一个对象，支持合并单元格处理。
- **腾讯云 COS 集成**：自动将 Excel 中嵌入的文件上传到腾讯云 COS，并在 JSON 输出中使用 COS URL。
- **高性能转换**：
    - **虚拟线程并发**：利用 Java 21 虚拟线程 (Virtual Threads) 技术并行上传内嵌文件，大幅提升包含大量图片/文件的 Excel 处理速度。
    - **资源优化**：重构解析逻辑，确保 Workbook 只打开一次，减少重复 I/O，显著降低大文件转换耗时。
- **会话管理**：
    - 提供 `openFile`、`foreach`、`reset` 接口，支持大文件分批次读取。
    - 会话状态（Session）在内存中保持 24 小时，支持断点续传和指针管理。
- **缓存机制**：基于 MD5 的缓存机制，避免重复解析相同文件。
- **MCP SSE 支持**：提供服务器发送事件（SSE）端点用于 MCP 通信。

## 系统要求

- **Java 21** 或更高版本（必须，用于支持虚拟线程）
- Maven 3.6 或更高版本
- 腾讯云 COS 账户（用于文件存储）

## 安装与编译

### 1. 克隆项目
```bash
git clone <项目地址>
cd mysheet-mcp
```

### 2. 安装依赖
```bash
./mvnw clean install
```

### 3. 编译项目
```bash
./mvnw compile
```

### 4. 运行测试
```bash
./mvnw test
```

## 配置说明

配置文件位于 `src/main/resources/application.yml`，需要根据您的环境进行配置：

```yaml
# 服务器配置
server:
  port: 8080  # MCP SSE 服务器端口，默认为 8080

# Spring AI MCP 配置
spring:
  ai:
    mcp:
      server:
        sse-message-endpoint: /sse  # SSE 消息端点路径

# 存储配置
storage:
  file: /var/mysheet-mcp/upload  # 上传的 Excel 文件存储目录
  cache: /var/mysheet-mcp/cache   # 缓存 JSON 结果的目录

# 腾讯云 COS 配置
cos:
  secret-id: YOUR_SECRET_ID       # 腾讯云访问密钥 ID
  secret-key: YOUR_SECRET_KEY     # 腾讯云访问密钥 Key
  region: ap-beijing              # COS 存储桶地域
  bucket-name: YOUR_BUCKET_NAME   # COS 存储桶名称
```

## 启动服务器

### 1. 开发环境启动
```bash
./mvnw spring-boot:run
```

### 2. 生产环境启动
首先打包项目：
```bash
./mvnw clean package
```

然后运行生成的 JAR 文件：
```bash
java -jar target/mysheet-mcp-0.0.1-SNAPSHOT.jar
```

### 3. 验证服务器运行
服务器启动后，可以通过以下方式验证：
- 访问 `http://localhost:8080/actuator/health` 检查健康状态
- 查看控制台日志，确认 MCP 服务器已启动

## 使用方式 (MCP Tools)

本服务器提供以下 MCP 工具供 Agent 调用：

### 1. excel2Json
一次性将整个 Excel 文件转换为 JSON。

- **参数**：
    - `excelFileURL` (String): Excel 文件的 URL 或本地路径。
    - `type` (String): 解析模式，`basic` (默认) 或 `row-object`。
- **返回**：完整的 JSON 数据字符串。

### 2. openFile (会话模式)
打开 Excel 文件并创建一个读取会话，适用于大文件处理。

- **参数**：
    - `excelFileURL` (String): Excel 文件的 URL 或本地路径。
    - `type` (String): 解析模式，`basic` 或 `row-object`。
    - `offset` (int, 可选): 起始行号，默认为 0。
- **返回**：包含会话信息的 JSON 对象。
    ```json
    {
      "text": "sessionId_string",
      "sessionId": "sessionId_string",
      "files": [],
      "json": [ { "data": [] } ]
    }
    ```
    - `sessionId`: 会话唯一标识，用于后续操作。
    - `text`: 兼容性字段，同 `sessionId`。

### 3. foreach (遍历会话)
读取当前会话的下一批数据。

- **参数**：
    - `sessionId` (String): `openFile` 返回的会话 ID。
    - `limit` (int, 可选): 读取行数限制，默认为 10。
- **返回**：包含数据的 JSON 字符串。
    - 如果还有数据：返回数据 JSON。
    - 如果已读完：返回提示信息或空数据。

### 4. reset (重置会话)
重置会话的读取指针到指定位置。

- **参数**：
    - `sessionId` (String): 会话 ID。
    - `offset` (int): 重置到的行号（从 0 开始）。
- **返回**：重置成功的提示信息。

## 解析模式详解

### 基础模式（Basic Mode）
默认解析模式，返回简化的 JSON 结构：
- 每个单元格直接返回其值
- 适合简单的数据提取需求

### 行对象模式（Row-Object Mode）
使用 `type="row-object"` 参数启用：
- 每一行转换为一个 JSON 对象
- 表头行定义属性名称
- 合并单元格通过重复值处理
- 每个单元格包含 `type` 和 `value` 字段

**单元格类型说明：**
- `text`：文本类型
- `number`：数字类型
- `money`：货币类型（检测到货币符号时自动识别）
- `date`：日期类型，格式化为 `yyyy-MM-dd`
- `file`：文件类型，值为 COS URL
- `boolean`：布尔类型

## 典型输入输出示例

### 输入 Excel 文件结构
假设有一个 `test.xls` 文件，包含以下内容：
- 表头：A1="", B1="标题1", C1="标题2", D1="标题1"
- 数据行包含文本、数字、日期、文件等不同类型数据

### 输出 JSON 示例（Row-Object 模式）

```json
{
  "header": {
    "A1": "",
    "B1": "标题1",
    "C1": "标题2",
    "D1": "标题1"
  },
  "data": [
    {
      "index": 1,
      "A1": {
        "type": "text",
        "value": "属性1"
      },
      "B1": {
        "type": "number",
        "value": 1
      },
      "C1": {
        "type": "date",
        "value": "2025-05-18"
      },
      "D1": {
        "type": "file",
        "value": "https://chinamobile-1320739042.cos.ap-beijing.myqcloud.com/attachment/20251230001822713_test.png"
      }
    }
  ],
  "filename": "test.xls",
  "md5": "a1b2c3d4e5f678901234567890123456"
}
```

## 缓存机制

### 工作原理
1. 系统计算上传 Excel 文件的 MD5 哈希值
2. 检查缓存目录中是否存在对应的 JSON 文件（文件名格式：`{md5}_{type}.json`）
3. 如果缓存存在且有效，直接返回缓存内容
4. 如果缓存不存在或无效，重新解析文件并保存到缓存

### 会话缓存
- 会话数据（SessionData）存储在内存中（TimedCache）。
- 有效期：**24 小时**。
- 清理策略：每小时自动清理过期会话。

## 故障排除

### 常见问题

1. **服务器启动失败**
   - **检查 Java 版本是否为 21+** (本项目使用了虚拟线程，必须使用 JDK 21)。
   - 检查端口 8080 是否被占用
   - 检查 COS 配置是否正确

2. **文件上传失败**
   - 检查 COS 密钥是否有上传权限
   - 检查存储桶地域是否正确
   - 检查网络连接

3. **解析错误**
   - 检查 Excel 文件格式是否支持（支持 .xls 和 .xlsx）
   - 检查文件是否损坏
   - 查看服务器日志获取详细错误信息

### 日志查看
```bash
# 查看实时日志
tail -f logs/application.log

# 或查看控制台输出
```

## 项目结构

```
mysheet-mcp/
├── src/
│   ├── main/
│   │   ├── java/link/wo/mysheetmcp/
│   │   │   ├── service/          # 服务层
│   │   │   │   ├── Excel2JsonService.java # 核心业务逻辑，含会话管理
│   │   │   │   └── CosService.java
│   │   │   ├── util/             # 工具类
│   │   │   │   └── Excel2JsonUtil.java    # Excel 解析与虚拟线程优化
│   │   │   └── McpServerApplication.java
│   │   └── resources/
│   │       └── application.yml   # 配置文件
│   └── test/                     # 测试代码
├── pom.xml                       # Maven 配置 (需 Java 21)
└── README.md                     # 本文档
```
