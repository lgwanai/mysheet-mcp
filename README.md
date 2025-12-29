# mysheet MCP 服务器

Mysheet-mcp 是一个用于将 Excel 文件转换为 JSON 格式的 MCP 服务器，特别适用于在大型模型或智能体中进行数据处理。本项目支持解析 Excel 文件，并将 Excel 中嵌入的文件自动上传到腾讯云 COS（对象存储）。

## 主要功能

- **Excel 转 JSON**：将上传的 Excel 文件转换为 JSON 格式，包含详细的单元格类型信息（文本、数字、货币、日期等）。
- **行对象模式**：支持 "row-object" 解析模式，将每一行视为一个对象，支持合并单元格处理。
- **腾讯云 COS 集成**：自动将 Excel 中嵌入的文件上传到腾讯云 COS，并在 JSON 输出中使用 COS URL。
- **缓存机制**：基于 MD5 的缓存机制，避免重复解析相同文件。
- **MCP SSE 支持**：提供服务器发送事件（SSE）端点用于 MCP 通信。

## 系统要求

- Java 17 或更高版本
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

### 配置参数详解

#### 服务器配置
- `server.port`：MCP 服务器监听的端口号，默认为 8080。
- `spring.ai.mcp.server.sse-message-endpoint`：SSE 消息端点路径，客户端通过此路径连接 MCP 服务器。

#### 存储配置
- `storage.file`：上传 Excel 文件临时存储目录，确保该目录有读写权限。
- `storage.cache`：缓存目录，用于存储已解析文件的 JSON 结果，基于 MD5 值进行缓存。

#### 腾讯云 COS 配置
- `cos.secret-id`：腾讯云访问密钥 ID，从腾讯云控制台获取。
- `cos.secret-key`：腾讯云访问密钥 Key，从腾讯云控制台获取。
- `cos.region`：COS 存储桶所在地域，如 `ap-beijing`（北京）。
- `cos.bucket-name`：COS 存储桶名称，格式为 `bucketname-appid`。

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

## 使用方式

### 1. 通过 MCP 客户端连接
MCP 客户端可以通过 SSE 端点连接到服务器：
```
http://localhost:8080/sse
```

### 2. 调用 Excel 转 JSON 服务
通过 MCP 调用 `excel2Json` 工具：

**参数说明：**
- `excelFileURL`：Excel 文件的 URL 地址（必需）
- `type`：解析模式，可选值：
  - `basic`：基础模式（默认）
  - `row-object`：行对象模式

**示例调用：**
```json
{
  "tool": "excel2Json",
  "parameters": {
    "excelFileURL": "https://example.com/test.xlsx",
    "type": "row-object"
  }
}
```

### 3. 直接调用 API（可选）
服务器也提供 REST API 端点：
```bash
curl -X POST http://localhost:8080/api/excel2json \
  -H "Content-Type: application/json" \
  -d '{"url": "https://example.com/test.xlsx", "type": "row-object"}'
```

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
    },
    {
      "index": 2,
      "A1": {
        "type": "text",
        "value": "属性2"
      },
      "B1": {
        "type": "money",
        "value": "￥1,000.00"
      },
      "C1": {
        "type": "money",
        "value": "￥1,000.00"
      },
      "D1": {
        "type": "number",
        "value": 0.5
      }
    }
  ],
  "filename": "test.xls",
  "md5": "a1b2c3d4e5f678901234567890123456"
}
```

### 输出字段说明
- `header`：表头信息，键为单元格位置，值为表头文本
- `data`：数据行数组，每行包含：
  - `index`：行索引（从1开始）
  - 单元格数据：键为单元格位置，值为包含 `type` 和 `value` 的对象
- `filename`：原始文件名
- `md5`：文件的 MD5 哈希值，用于缓存标识

## 缓存机制

### 工作原理
1. 系统计算上传 Excel 文件的 MD5 哈希值
2. 检查缓存目录中是否存在对应的 JSON 文件（文件名格式：`{md5}_{type}.json`）
3. 如果缓存存在且有效，直接返回缓存内容
4. 如果缓存不存在或无效，重新解析文件并保存到缓存

### 缓存文件命名规则
- 基础模式：`{md5}.json`
- 行对象模式：`{md5}_row-object.json`

### 手动清理缓存
```bash
# 清理所有缓存
rm -rf /var/mysheet-mcp/cache/*

# 或通过配置的 storage.cache 目录清理
```

## 故障排除

### 常见问题

1. **服务器启动失败**
   - 检查 Java 版本是否为 17+
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
│   │   │   │   ├── Excel2JsonService.java
│   │   │   │   └── CosService.java
│   │   │   ├── util/             # 工具类
│   │   │   │   └── Excel2JsonUtil.java
│   │   │   └── McpServerApplication.java
│   │   └── resources/
│   │       └── application.yml   # 配置文件
│   └── test/                     # 测试代码
├── pom.xml                       # Maven 配置
└── README.md                     # 本文档
```

## 贡献指南

1. Fork 项目
2. 创建功能分支
3. 提交更改
4. 推送到分支
5. 创建 Pull Request

## 许可证

[在此添加许可证信息]

## 支持与反馈

如有问题或建议，请通过以下方式联系：
- 提交 Issue
- 发送邮件
- 其他联系方式
