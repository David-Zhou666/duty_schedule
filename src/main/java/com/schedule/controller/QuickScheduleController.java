package com.schedule.controller;

import com.schedule.model.DutySchedule;
import com.schedule.model.Person;
import com.schedule.service.ScheduleService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import java.io.*;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

/**
 * 快速排班控制器
 * 用于从预设的Excel文件直接生成排班表
 */
@Controller
public class QuickScheduleController {
    private static final Logger logger = LoggerFactory.getLogger(QuickScheduleController.class);
    
    @Autowired
    private ScheduleService scheduleService;
    
    /**
     * 从duty_temple.xlsx直接生成排班表并下载
     */
    @GetMapping("/download-schedule")
    public ResponseEntity<byte[]> downloadSchedule() {
        try {
            logger.info("开始从duty_temple.xlsx生成排班表");
            
            String filePath = "src/main/resources/duty_temple.xlsx";
            List<Person> persons = readExcelFile(filePath);
            
            if (persons.isEmpty()) {
                logger.warn("未读取到任何人员信息");
                return ResponseEntity.badRequest().build();
            }
            
            logger.info("成功读取 {} 名人员", persons.size());
            
            // 生成排班表
            List<DutySchedule> schedules = scheduleService.generateWeeklySchedule(persons);
            
            // 导出Excel
            byte[] excelData = scheduleService.createWeeklyScheduleWorkbook(schedules);
            
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData("attachment", "duty_schedule_result.xlsx");
            
            logger.info("排班表导出成功");
            return ResponseEntity.ok()
                    .headers(headers)
                    .body(excelData);
            
        } catch (Exception e) {
            logger.error("生成排班表失败", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }
    
    /**
     * 从Excel文件读取人员信息
     */
    private List<Person> readExcelFile(String filePath) throws IOException {
        List<Person> persons = new ArrayList<>();
        
        File file = new File(filePath);
        if (!file.exists()) {
            logger.warn("文件不存在: {}", filePath);
            // 尝试使用classpath路径
            filePath = "target/classes/" + filePath;
            file = new File(filePath);
            if (!file.exists()) {
                logger.error("无法找到文件");
                return persons;
            }
        }
        
        try (FileInputStream fis = new FileInputStream(file)) {
            org.apache.poi.ss.usermodel.Workbook workbook = 
                    new org.apache.poi.xssf.usermodel.XSSFWorkbook(fis);
            
            int sheetCount = workbook.getNumberOfSheets();
            logger.info("Excel文件Sheet页数: {}", sheetCount);
            
            for (int sheetIdx = 0; sheetIdx < sheetCount; sheetIdx++) {
                org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(sheetIdx);
                String sheetName = sheet.getSheetName();
                logger.info("正在读取 Sheet: {}", sheetName);
                
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
    private Person parseRow(org.apache.poi.ss.usermodel.Row row, boolean isMale) {
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
            logger.error("解析行出错: {}", e.getMessage());
            return null;
        }
    }
    
    /**
     * 获取单元格字符串值
     */
    private String getCellStringValue(org.apache.poi.ss.usermodel.Cell cell) {
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
}
