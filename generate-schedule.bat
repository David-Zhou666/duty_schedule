@echo off
REM 排班表生成脚本
REM 用法: generate-schedule.bat

echo.
echo ========================================
echo   值日排班表生成器
echo ========================================
echo.

REM 使用Maven运行ScheduleMain
echo 正在生成排班表...
echo.

call .\apache-maven-3.9.6\bin\mvn.cmd clean compile 2>nul

if exist target\classes (
    echo 编译成功
    echo 正在从 duty_temple.xlsx 读取数据...
    echo.
    
    REM 由于难以直接运行Java程序（依赖问题），我们使用Maven exec插件
    call .\apache-maven-3.9.6\bin\mvn.cmd exec:java -Dexec.mainClass="com.schedule.ScheduleMain" -DskipTests 2>nul
) else (
    echo 编译失败
    exit /b 1
)

echo.
echo ========================================
echo 排班表已生成！
echo 文件位置: duty_schedule_result.xlsx
echo ========================================
echo.

pause
