package com.schedule;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 值日排班生成器
 * 用于读取duty_temple.xlsx并生成排班表
 */
public class DutyScheduleGenerator {
    
    private static final int MALE_PER_DAY = 4;
    private static final int FEMALE_PER_DAY = 2;
    private static final int WEEKLY_LIMIT = 5; // 每周最多值日5次
    private static final int DAYS_PER_WEEK = 5; // 周一到周五

    private List<Student> maleStudents;
    private List<Student> femaleStudents;

    public DutyScheduleGenerator() {
        this.maleStudents = new ArrayList<>();
        this.femaleStudents = new ArrayList<>();
    }

    /**
     * 读取duty_temple.xlsx文件
     */
    public void readExcelFile(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            System.out.println("开始读取Excel文件: " + filePath);
            System.out.println("Sheet页数: " + workbook.getNumberOfSheets());
            
            for (int sheetIdx = 0; sheetIdx < workbook.getNumberOfSheets(); sheetIdx++) {
                Sheet sheet = workbook.getSheetAt(sheetIdx);
                String sheetName = sheet.getSheetName();
                System.out.println("\n正在读取 Sheet: " + sheetName);
                
                readSheet(sheet, sheetName);
            }
        }
    }

    /**
     * 读取单个Sheet
     */
    private void readSheet(Sheet sheet, String sheetName) {
        // 判断是男生还是女生（根据sheet名称或其他方式）
        boolean isMale = sheetName.contains("男");
        
        // 跳过表头（第一行）
        for (int rowIdx = 1; rowIdx <= sheet.getLastRowNum(); rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row == null) continue;
            
            // 读取姓名（第一列）
            String name = getCellStringValue(row.getCell(0));
            if (name == null || name.trim().isEmpty()) continue;
            
            // 读取周一到周五的可用性
            boolean[] availability = new boolean[DAYS_PER_WEEK];
            for (int dayIdx = 0; dayIdx < DAYS_PER_WEEK; dayIdx++) {
                String value = getCellStringValue(row.getCell(1 + dayIdx));
                availability[dayIdx] = "是".equals(value) || "是".equals(value);
            }
            
            Student student = new Student(name, isMale, availability);
            if (isMale) {
                maleStudents.add(student);
                System.out.println("  添加男生: " + name);
            } else {
                femaleStudents.add(student);
                System.out.println("  添加女生: " + name);
            }
        }
    }

    /**
     * 生成5天排班表
     */
    public List<DaySchedule> generateWeeklySchedule() {
        System.out.println("\n开始生成排班表...");
        System.out.println("男生数: " + maleStudents.size() + ", 女生数: " + femaleStudents.size());
        
        List<DaySchedule> weekSchedule = new ArrayList<>();
        String[] dayNames = {"周一", "周二", "周三", "周四", "周五"};
        
        for (int dayIdx = 0; dayIdx < DAYS_PER_WEEK; dayIdx++) {
            List<String> maleAssignments = selectStudents(maleStudents, MALE_PER_DAY, dayIdx);
            List<String> femaleAssignments = selectStudents(femaleStudents, FEMALE_PER_DAY, dayIdx);
            
            DaySchedule schedule = new DaySchedule(dayNames[dayIdx], maleAssignments, femaleAssignments);
            weekSchedule.add(schedule);
            
            System.out.println(dayNames[dayIdx] + ": 男生=[" + String.join(", ", maleAssignments) + 
                             "], 女生=[" + String.join(", ", femaleAssignments) + "]");
        }
        
        return weekSchedule;
    }

    /**
     * 为某一天选择学生
     * 使用贪心算法：优先选择该天可用且本周排班次数最少的学生
     */
    private List<String> selectStudents(List<Student> students, int count, int dayIdx) {
        List<String> selected = new ArrayList<>();
        
        // 过滤出该天可用且未超过周限制的学生
        List<Student> available = students.stream()
                .filter(s -> s.isAvailableOn(dayIdx) && s.getWeeklyCount() < WEEKLY_LIMIT)
                .sorted(Comparator.comparingInt(Student::getWeeklyCount))
                .collect(Collectors.toList());
        
        // 选择前count个
        for (int i = 0; i < Math.min(count, available.size()); i++) {
            Student student = available.get(i);
            selected.add(student.getName());
            student.incrementWeekly();
        }
        
        // 如果不足，放宽约束（允许超过限制）
        if (selected.size() < count) {
            List<Student> remaining = students.stream()
                    .filter(s -> s.isAvailableOn(dayIdx) && !selected.contains(s.getName()))
                    .sorted(Comparator.comparingInt(Student::getWeeklyCount))
                    .collect(Collectors.toList());
            
            for (int i = 0; i < Math.min(count - selected.size(), remaining.size()); i++) {
                Student student = remaining.get(i);
                selected.add(student.getName());
                student.incrementWeekly();
            }
        }
        
        return selected;
    }

    /**
     * 导出排班表到Excel
     */
    public void exportToExcel(List<DaySchedule> schedules, String outputPath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("排班表");
            
            // 创建表头
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("周次");
            headerRow.createCell(1).setCellValue("男生值班(4人)");
            headerRow.createCell(2).setCellValue("女生值班(2人)");
            
            // 填充数据
            int rowIdx = 1;
            for (DaySchedule schedule : schedules) {
                Row row = sheet.createRow(rowIdx++);
                row.createCell(0).setCellValue(schedule.getDay());
                row.createCell(1).setCellValue(String.join(", ", schedule.getMaleAssignments()));
                row.createCell(2).setCellValue(String.join(", ", schedule.getFemaleAssignments()));
            }
            
            // 自动调整列宽
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            sheet.autoSizeColumn(2);
            
            // 写入文件
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                workbook.write(fos);
                System.out.println("\n排班表已导出到: " + outputPath);
            }
        }
    }

    /**
     * 获取单元格字符串值
     */
    private String getCellStringValue(Cell cell) {
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

    public static void main(String[] args) throws IOException {
        String inputFile = "src/main/resources/duty_temple.xlsx";
        String outputFile = "duty_schedule_output.xlsx";
        
        DutyScheduleGenerator generator = new DutyScheduleGenerator();
        generator.readExcelFile(inputFile);
        List<DaySchedule> schedules = generator.generateWeeklySchedule();
        generator.exportToExcel(schedules, outputFile);
        
        System.out.println("\n排班完成！");
    }

    /**
     * 学生类
     */
    static class Student {
        private String name;
        private boolean isMale;
        private boolean[] availability; // 五天的可用性
        private int weeklyCount;

        public Student(String name, boolean isMale, boolean[] availability) {
            this.name = name;
            this.isMale = isMale;
            this.availability = availability;
            this.weeklyCount = 0;
        }

        public String getName() {
            return name;
        }

        public boolean isMale() {
            return isMale;
        }

        public boolean isAvailableOn(int dayIdx) {
            return dayIdx >= 0 && dayIdx < availability.length && availability[dayIdx];
        }

        public int getWeeklyCount() {
            return weeklyCount;
        }

        public void incrementWeekly() {
            this.weeklyCount++;
        }

        public void resetWeekly() {
            this.weeklyCount = 0;
        }
    }

    /**
     * 每天的排班信息
     */
    static class DaySchedule {
        private String day;
        private List<String> maleAssignments;
        private List<String> femaleAssignments;

        public DaySchedule(String day, List<String> maleAssignments, List<String> femaleAssignments) {
            this.day = day;
            this.maleAssignments = maleAssignments;
            this.femaleAssignments = femaleAssignments;
        }

        public String getDay() {
            return day;
        }

        public List<String> getMaleAssignments() {
            return maleAssignments;
        }

        public List<String> getFemaleAssignments() {
            return femaleAssignments;
        }
    }
}
