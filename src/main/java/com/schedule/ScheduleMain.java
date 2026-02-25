package com.schedule;

import com.schedule.model.DutySchedule;
import com.schedule.model.Person;
import com.schedule.service.ScheduleService;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * 排班主程序
 * 用于直接读取duty_temple.xlsx并生成排班表
 */
public class ScheduleMain {

    public static void main(String[] args) throws IOException {
        System.out.println("开始生成排班表...");
        
        String inputFile = "src/main/resources/duty_temple.xlsx";
        String outputFile = "duty_schedule_result.xlsx";
        
        // 创建一个简单的MultipartFile包装器来读取本地文件
        File file = new File(inputFile);
        if (!file.exists()) {
            System.err.println("输入文件不存在: " + inputFile);
            System.exit(1);
        }
        
        try {
            // 使用现有的services
            ScheduleService scheduleService = new ScheduleService();
            
            // 读取Excel文件
            System.out.println("正在读取 " + inputFile);
            List<Person> persons = readExcelFile(file);
            
            if (persons.isEmpty()) {
                System.err.println("未读取到任何人员信息");
                System.exit(1);
            }
            
            System.out.println("成功读取 " + persons.size() + " 名人员");
            printPersonInfo(persons);
            
            // 生成排班表
            List<DutySchedule> schedules = scheduleService.generateWeeklySchedule(persons);
            
            // 导出Excel
            System.out.println("正在生成Excel文件...");
            byte[] excelData = scheduleService.createWeeklyScheduleWorkbook(schedules);
            
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                fos.write(excelData);
                System.out.println("排班表已成功导出到: " + outputFile);
            }
            
            System.out.println("\n排班完成！详细信息如下：");
            printScheduleInfo(schedules);
            
        } catch (Exception e) {
            System.err.println("发生错误: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
    
    /**
     * 从Excel文件读取人员信息
     */
    private static List<Person> readExcelFile(File file) throws IOException {
        List<Person> persons = new java.util.ArrayList<>();
        
        try (FileInputStream fis = new FileInputStream(file)) {
            org.apache.poi.ss.usermodel.Workbook workbook = 
                    new org.apache.poi.xssf.usermodel.XSSFWorkbook(fis);
            
            int sheetCount = workbook.getNumberOfSheets();
            System.out.println("Sheet页数: " + sheetCount);
            
            for (int sheetIdx = 0; sheetIdx < sheetCount; sheetIdx++) {
                org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(sheetIdx);
                String sheetName = sheet.getSheetName();
                System.out.println("正在读取 Sheet: " + sheetName);
                
                // 判断性别（根据sheet名称）
                boolean isMale = sheetName.contains("男");
                
                // 从第二行开始（第一行是表头）
                for (int rowIdx = 1; rowIdx <= sheet.getLastRowNum(); rowIdx++) {
                    org.apache.poi.ss.usermodel.Row row = sheet.getRow(rowIdx);
                    if (row == null) continue;
                    
                    Person person = parseRow(row, isMale);
                    if (person != null) {
                        persons.add(person);
                    }
                }
            }
            
            workbook.close();
        }
        
        return persons;
    }
    
    /**
     * 解析Excel行
     */
    private static Person parseRow(org.apache.poi.ss.usermodel.Row row, boolean isMale) {
        try {
            org.apache.poi.ss.usermodel.Cell nameCell = row.getCell(0);
            String name = getCellStringValue(nameCell);
            
            if (name == null || name.trim().isEmpty()) {
                return null;
            }
            
            name = name.trim();
            String gender = isMale ? "男" : "女";
            
            Person person = new Person(name, gender, false);
            
            // 读取周一到周五的可用性（列1-5）
            boolean[] availability = new boolean[5];
            for (int dayIdx = 0; dayIdx < 5; dayIdx++) {
                org.apache.poi.ss.usermodel.Cell cell = row.getCell(dayIdx + 1);
                String value = getCellStringValue(cell);
                
                // "是"或其他值视为可用，空值也视为可用
                boolean available = value == null || value.isEmpty() || "是".equals(value);
                availability[dayIdx] = available;
            }
            
            person.setAvailability(availability);
            return person;
            
        } catch (Exception e) {
            System.err.println("解析行出错: " + e.getMessage());
            return null;
        }
    }
    
    /**
     * 获取单元格字符串值
     */
    private static String getCellStringValue(org.apache.poi.ss.usermodel.Cell cell) {
        if (cell == null) return null;
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return cell.getBooleanCellValue() ? "是" : "否";
            default:
                return null;
        }
    }
    
    /**
     * 打印人员信息统计
     */
    private static void printPersonInfo(List<Person> persons) {
        long maleCount = persons.stream().filter(p -> "男".equals(p.getGender())).count();
        long femaleCount = persons.stream().filter(p -> "女".equals(p.getGender())).count();
        
        System.out.println("男生: " + maleCount + "人");
        System.out.println("女生: " + femaleCount + "人");
        System.out.println();
    }
    
    /**
     * 打印排班信息
     */
    private static void printScheduleInfo(List<DutySchedule> schedules) {
        String[] dayNames = {"周一", "周二", "周三", "周四", "周五"};
        
        System.out.println("\n" + "=".repeat(80));
        System.out.println("排班表详情");
        System.out.println("=".repeat(80));
        
        for (int i = 0; i < schedules.size(); i++) {
            DutySchedule schedule = schedules.get(i);
            
            String males = schedule.getMaleDuty().stream()
                    .map(Person::getName)
                    .reduce((a, b) -> a + "、" + b)
                    .orElse("-");
            
            String females = schedule.getFemaleDuty().stream()
                    .map(Person::getName)
                    .reduce((a, b) -> a + "、" + b)
                    .orElse("-");
            
            System.out.println(String.format("%-8s | 男生(4人): %-30s | 女生(2人): %s", 
                    dayNames[i], males, females));
        }
        
        System.out.println("=".repeat(80));
    }
}
