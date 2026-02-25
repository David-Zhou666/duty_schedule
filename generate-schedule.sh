#!/bin/bash
# 排班表生成脚本
# 用法: ./generate-schedule.sh

cd "$(dirname "$0")"

CLASSPATH="target/classes"
for jar in target/lib/*.jar; do
    CLASSPATH="$CLASSPATH:$jar"
done

java -cp "$CLASSPATH:poi-5.2.3.jar:poi-ooxml-5.2.3.jar:poi-ooxml-lite-5.2.3.jar:xmlbeans-5.1.1.jar" \
    com.schedule.ScheduleMain

if [ $? -eq 0 ]; then
    echo ""
    echo "排班表已生成: duty_schedule_result.xlsx"
    echo "请查看输出目录中的Excel文件"
fi
