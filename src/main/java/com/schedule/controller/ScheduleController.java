package com.schedule.controller;

import com.schedule.model.DutySchedule;
import com.schedule.model.Person;
import com.schedule.service.ExcelParserService;
import com.schedule.service.ScheduleService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import java.io.IOException;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * 排班控制器
 * darkzhou 
 */
@Controller
public class ScheduleController {
    private static final Logger logger = LoggerFactory.getLogger(ScheduleController.class);

    @Autowired
    private ExcelParserService excelParserService;

    @Autowired
    private ScheduleService scheduleService;

    /**
     * 首页
     */
    @GetMapping("/")
    public String index() {
        return "index";
    }

    /**
     * 上传 Excel 并生成排班
     */
    @PostMapping("/upload")
    @ResponseBody
    public Map<String, Object> uploadAndSchedule(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "startDate", required = false) String startDate,
            @RequestParam(value = "endDate", required = false) String endDate,
            @RequestParam(value = "holidays", required = false) String holidays) {

        Map<String, Object> result = new HashMap<>();

        try {
            // 1. 解析 Excel 文件
            if (file.isEmpty()) {
                result.put("success", false);
                result.put("message", "文件不能为空");
                return result;
            }

            logger.info("接收到文件: {}, 大小: {} bytes", file.getOriginalFilename(), file.getSize());

            List<Person> persons = excelParserService.parseExcelFile(file);

            if (persons.isEmpty()) {
                result.put("success", false);
                result.put("message", "未解析到任何人员信息，请检查 Excel 格式");
                return result;
            }

            logger.info("成功解析 {} 名人员", persons.size());

            // 2. 确定日期范围（默认本月）
            LocalDate start = startDate != null && !startDate.isEmpty() ?
                    LocalDate.parse(startDate) : LocalDate.now().withDayOfMonth(1);
            LocalDate end = endDate != null && !endDate.isEmpty() ?
                    LocalDate.parse(endDate) : start.withDayOfMonth(start.lengthOfMonth());

            logger.info("排班日期范围: {} 到 {}", start, end);

            // 3. 生成节假日列表
            Set<LocalDate> holidaySet = scheduleService.generateHolidays2026();

            // 4. 生成排班表
            List<DutySchedule> schedules = scheduleService.generateSchedule(persons, start, end, holidaySet);

            // 5. 格式化排班表
            String scheduleTable = scheduleService.formatScheduleTable(schedules);

            result.put("success", true);
            result.put("message", "排班成功");
            result.put("personCount", persons.size());
            result.put("scheduleCount", schedules.size());
            result.put("scheduleTable", scheduleTable);

            logger.info("排班完成，共生成 {} 条记录", schedules.size());

        } catch (IOException e) {
            logger.error("文件解析失败", e);
            result.put("success", false);
            result.put("message", "文件解析失败: " + e.getMessage());
        } catch (Exception e) {
            logger.error("排班失败", e);
            result.put("success", false);
            result.put("message", "排班失败: " + e.getMessage());
        }

        return result;
    }

    /**
     * 下载排班表（纯文本格式）
     */
    @PostMapping("/download")
    @ResponseBody
    public String downloadSchedule(@RequestBody Map<String, Object> params) {
        String scheduleTable = (String) params.get("scheduleTable");
        return scheduleTable;
    }

    /**
     * 获取节假日列表
     */
    @GetMapping("/holidays")
    @ResponseBody
    public Set<LocalDate> getHolidays() {
        return scheduleService.generateHolidays2026();
    }

    /**
     * 上传并直接导出排班表为 Excel 文件
     */
    @PostMapping(value = "/export", produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public ResponseEntity<byte[]> exportScheduleExcel(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "startDate", required = false) String startDate,
            @RequestParam(value = "endDate", required = false) String endDate) {

        try {
            if (file.isEmpty()) {
                return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(null);
            }

            List<Person> persons = excelParserService.parseExcelFile(file);
            if (persons.isEmpty()) {
                return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(null);
            }

            LocalDate start = startDate != null && !startDate.isEmpty() ?
                    LocalDate.parse(startDate) : LocalDate.now().withDayOfMonth(1);
            LocalDate end = endDate != null && !endDate.isEmpty() ?
                    LocalDate.parse(endDate) : start.withDayOfMonth(start.lengthOfMonth());

            Set<LocalDate> holidaySet = scheduleService.generateHolidays2026();
            List<com.schedule.model.DutySchedule> schedules = scheduleService.generateSchedule(persons, start, end, holidaySet);

            byte[] excelBytes = scheduleService.createScheduleWorkbook(schedules);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"duty_schedule.xlsx\"");
            headers.setContentLength(excelBytes.length);

            return new ResponseEntity<>(excelBytes, headers, HttpStatus.OK);

        } catch (IOException e) {
            logger.error("导出失败", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
}
