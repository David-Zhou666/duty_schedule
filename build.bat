@echo off
echo Building Duty Scheduler...
echo.

REM 检查 Maven 是否安装
mvn -version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Maven is not installed or not in PATH
    echo Please install Maven from https://maven.apache.org/download.cgi
    exit /b 1
)

REM 编译打包
echo Running Maven clean package...
call mvn clean package

if %errorlevel% equ 0 (
    echo.
    echo ================================================
    echo Build successful!
    echo.
    echo WAR file location: target\duty-scheduler.war
    echo.
    echo To deploy to Tomcat:
    echo 1. Copy target\duty-scheduler.war to %TOMCAT_HOME%\webapps\
    echo 2. Start Tomcat
    echo 3. Access: http://localhost:8080/duty-scheduler/
    echo ================================================
) else (
    echo.
    echo Build failed! Please check the errors above.
)

pause
