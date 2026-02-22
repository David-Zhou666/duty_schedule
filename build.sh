#!/bin/bash

echo "Building Duty Scheduler..."
echo ""

# Check if Maven is installed
if ! command -v mvn &> /dev/null; then
    echo "Error: Maven is not installed or not in PATH"
    echo "Please install Maven from https://maven.apache.org/download.cgi"
    exit 1
fi

# Build the project
echo "Running Maven clean package..."
mvn clean package

if [ $? -eq 0 ]; then
    echo ""
    echo "==========================================="
    echo "Build successful!"
    echo ""
    echo "WAR file location: target/duty-scheduler.war"
    echo ""
    echo "To deploy to Tomcat:"
    echo "1. Copy target/duty-scheduler.war to \$TOMCAT_HOME/webapps/"
    echo "2. Start Tomcat"
    echo "3. Access: http://localhost:8080/duty-scheduler/"
    echo "==========================================="
else
    echo ""
    echo "Build failed! Please check the errors above."
    exit 1
fi
