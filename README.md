# 值日排班系统

基于 Tomcat 的值日排班管理系统，支持 Excel 文件上传、自动排班、节假日调整等功能。

## 功能特性

- ✅ 支持 Excel 文件上传（.xlsx 格式）
- ✅ 支持多个 Sheet 页解析
- ✅ 智能排班算法
  - 每人一周不超过 2 次
  - 每人一个月不超过 4 次
  - 优先安排纪检委员
- ✅ 工作日排班（周一至周五）
  - 男生每天 4 人
  - 女生 2 人
- ✅ 节假日自动跳过
- ✅ 自动生成排班表格

## Excel 格式说明

Excel 文件应包含以下列：

| 列 | 说明 | 示例 |
|----|------|------|
| A | 姓名 | 张三 |
| B | 性别 | 男/女 |
| C | 是否纪检委员 | 是/否 |

示例数据：
```
张三  男  是
李四  女  否
王五  男  是
...
```

支持多个 Sheet 页，所有 Sheet 页都会被解析。

## 部署步骤

### 1. 环境要求

- JDK 1.8 或以上
- Apache Tomcat 8.5 或以上
- Maven 3.6 或以上

### 2. 编译打包

```bash
mvn clean package
```

打包成功后，会在 `target` 目录下生成 `duty-scheduler.war` 文件。

### 3. 部署到 Tomcat

将 `duty-scheduler.war` 复制到 Tomcat 的 `webapps` 目录：

```bash
cp target/duty-scheduler.war $TOMCAT_HOME/webapps/
```

### 4. 启动 Tomcat

```bash
$TOMCAT_HOME/bin/startup.sh  # Linux/Mac
$TOMCAT_HOME/bin/startup.bat # Windows
```

### 5. 访问系统

在浏览器中访问：`http://localhost:8080/duty-scheduler/`

## 项目结构

```
duty-scheduler/
├── src/
│   └── main/
│       ├── java/
│       │   └── com/schedule/
│       │       ├── controller/    # 控制器层
│       │       │   └── ScheduleController.java
│       │       ├── model/        # 数据模型
│       │       │   ├── Person.java
│       │       │   ├── DutySchedule.java
│       │       │   └── Holiday.java
│       │       └── service/      # 业务逻辑层
│       │           ├── ExcelParserService.java
│       │           └── ScheduleService.java
│       └── webapp/
│           ├── WEB-INF/
│           │   ├── web.xml
│           │   └── spring-mvc.xml
│           └── index.html        # 前端页面
├── pom.xml
└── README.md
```

## API 接口

### 上传 Excel 并生成排班

**URL:** `POST /upload`

**参数:**
- `file`: Excel 文件（必填）
- `startDate`: 开始日期（可选，格式：yyyy-MM-dd）
- `endDate`: 结束日期（可选，格式：yyyy-MM-dd）

**响应:**
```json
{
  "success": true,
  "message": "排班成功",
  "personCount": 20,
  "scheduleCount": 22,
  "scheduleTable": "排班表..."
}
```

### 获取节假日列表

**URL:** `GET /holidays`

**响应:** 返回 2026 年节假日列表

## 技术栈

- **后端框架:** Spring MVC 5.3.27
- **Excel 处理:** Apache POI 5.2.3
- **构建工具:** Maven
- **服务器:** Apache Tomcat

## 节假日配置

系统内置了 2026 年的中国法定节假日：

- 元旦、春节、清明节、劳动节
- 端午节、中秋节、国庆节

如果需要自定义节假日，请修改 `ScheduleService.java` 中的 `generateHolidays2026()` 方法。

## 许可证

MIT License
