package com.schedule.service;

import com.schedule.model.DutySchedule;
import com.schedule.model.Person;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.text.SimpleDateFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.temporal.TemporalAdjusters;
import java.util.*;
import java.util.stream.Collectors;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * 排班服务
 */
@Service
public class ScheduleService {
    private static final Logger logger = LoggerFactory.getLogger(ScheduleService.class);
    private static final int MALE_PER_DAY = 4;
    private static final int FEMALE_PER_DAY = 2;
    private static final int WEEKLY_LIMIT = 5; // 每人每周上班次数上限（改为5）
    private static final int MONTHLY_LIMIT = 20; // 每人每月上限（大数以避免过早限制）

    /**
     * 为指定日期范围生成排班表
     * @param persons 人员列表
     * @param startDate 开始日期
     * @param endDate 结束日期
     * @param holidays 节假日列表
     * @return 排班表
     */
    public List<DutySchedule> generateSchedule(List<Person> persons, LocalDate startDate,
                                               LocalDate endDate, Set<LocalDate> holidays) {
        logger.info("开始生成排班表: {} 到 {}", startDate, endDate);

        List<DutySchedule> schedules = new ArrayList<>();

        // 按性别和是否纪检委员分组
        Map<Boolean, List<Person>> maleGroups = persons.stream()
                .filter(p -> "男".equals(p.getGender()))
                .collect(Collectors.groupingBy(Person::isInspector));

        Map<Boolean, List<Person>> femaleGroups = persons.stream()
                .filter(p -> "女".equals(p.getGender()))
                .collect(Collectors.groupingBy(Person::isInspector));

        List<Person> maleInspectors = maleGroups.getOrDefault(true, new ArrayList<>());
        List<Person> maleNormal = maleGroups.getOrDefault(false, new ArrayList<>());
        List<Person> femaleInspectors = femaleGroups.getOrDefault(true, new ArrayList<>());
        List<Person> femaleNormal = femaleGroups.getOrDefault(false, new ArrayList<>());

        logger.info("人员统计 - 男生纪检委员: {}, 普通男生: {}, 女生纪检委员: {}, 普通女生: {}",
                maleInspectors.size(), maleNormal.size(),
                femaleInspectors.size(), femaleNormal.size());

        LocalDate currentDate = startDate;
        while (!currentDate.isAfter(endDate)) {
            // 检查是否为节假日
            boolean isHoliday = holidays.contains(currentDate);

            if (isHoliday) {
                logger.info("日期 {} 是节假日，不排班", currentDate);
                schedules.add(new DutySchedule(
                        java.sql.Date.valueOf(currentDate),
                        true,
                        new ArrayList<>(),
                        new ArrayList<>()
                ));
            } else if (currentDate.getDayOfWeek().getValue() <= 5) {
                // 周一到周五正常排班
                List<Person> maleDuty = selectPersons(maleInspectors, maleNormal, MALE_PER_DAY, currentDate.getDayOfWeek().getValue());
                List<Person> femaleDuty = selectPersons(femaleInspectors, femaleNormal, FEMALE_PER_DAY, currentDate.getDayOfWeek().getValue());

                DutySchedule schedule = new DutySchedule(
                        java.sql.Date.valueOf(currentDate),
                        false,
                        maleDuty,
                        femaleDuty
                );
                schedules.add(schedule);

                logger.info("{} 排班: 男生={}, 女生={}", currentDate,
                        maleDuty.stream().map(Person::getName).collect(Collectors.joining(",")),
                        femaleDuty.stream().map(Person::getName).collect(Collectors.joining(",")));
            }

            // 跨周时重置本周计数
            if (currentDate.getDayOfWeek() == DayOfWeek.SUNDAY) {
                resetWeeklyCount(persons);
            }

            currentDate = currentDate.plusDays(1);
        }

        logger.info("排班完成，共生成 {} 条记录", schedules.size());
        return schedules;
    }

    /**
     * 选择值班人员
     * 规则：优先排纪检委员，每人一周不超过2次，一个月不超过4次
     */
    private List<Person> selectPersons(List<Person> inspectors, List<Person> normal, int count, int dayOfWeekValue) {
        List<Person> selected = new ArrayList<>();
        List<Person> all = new ArrayList<>();

        // 过滤出当天可用的人员（dayOfWeekValue: 1=Mon .. 7=Sun）
        int weekdayIndex = dayOfWeekValue; // 1..7
        List<Person> availInspectors = new ArrayList<>();
        for (Person p : inspectors) {
            if (weekdayIndex >= 1 && weekdayIndex <= 5) {
                if (p.isAvailableOn(weekdayIndex)) availInspectors.add(p);
            }
        }
        List<Person> availNormal = new ArrayList<>();
        for (Person p : normal) {
            if (weekdayIndex >= 1 && weekdayIndex <= 5) {
                if (p.isAvailableOn(weekdayIndex)) availNormal.add(p);
            }
        }

        // 优先选择纪检委员（按已排班次数升序排序）
        availInspectors.sort(Comparator.comparingInt(Person::getWeeklyCount)
                .thenComparingInt(Person::getMonthlyCount));
        all.addAll(availInspectors);

        // 添加普通人员
        availNormal.sort(Comparator.comparingInt(Person::getWeeklyCount)
                .thenComparingInt(Person::getMonthlyCount));
        all.addAll(availNormal);

        // 选择符合条件的人员
        for (Person person : all) {
            if (selected.size() >= count) break;

            // 检查约束条件
            if (person.getWeeklyCount() < WEEKLY_LIMIT && person.getMonthlyCount() < MONTHLY_LIMIT) {
                selected.add(person);
                person.incrementWeekly();
                person.incrementMonthly();
                logger.debug("选择 {}: 本周{}, 本月{}", person.getName(),
                        person.getWeeklyCount(), person.getMonthlyCount());
            }
        }

        // 如果人数不足，放宽约束
        if (selected.size() < count) {
            logger.warn("可用人数不足，尝试放宽约束");
            for (Person person : all) {
                if (selected.size() >= count) break;
                if (!selected.contains(person)) {
                    selected.add(person);
                    person.incrementWeekly();
                    person.incrementMonthly();
                }
            }
        }

        return selected;
    }

    /**
     * 重置本周计数
     */
    private void resetWeeklyCount(List<Person> persons) {
        for (Person person : persons) {
            person.resetWeekly();
        }
        logger.info("新的一周开始，重置本周计数");
    }

    /**
     * 将排班表转换为表格字符串
     */
    public String formatScheduleTable(List<DutySchedule> schedules) {
        StringBuilder sb = new StringBuilder();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd E");

        sb.append("排班表\n");
        sb.append("=" + "=".repeat(80) + "\n");
        sb.append(String.format("%-12s %-8s %-30s %-20s\n",
                "日期", "是否节假日", "男生值班(4人)", "女生值班(2人)"));
        sb.append("-".repeat(80) + "\n");

        for (DutySchedule schedule : schedules) {
            String dateStr = sdf.format(schedule.getDate());
            String holidayStr = schedule.isHoliday() ? "是" : "否";

            String maleNames = schedule.getMaleDuty().stream()
                    .map(Person::getName)
                    .collect(Collectors.joining(","));

            String femaleNames = schedule.getFemaleDuty().stream()
                    .map(Person::getName)
                    .collect(Collectors.joining(","));

            if (schedule.isHoliday()) {
                maleNames = "-";
                femaleNames = "-";
            }

            sb.append(String.format("%-12s %-8s %-30s %-20s\n",
                    dateStr, holidayStr, maleNames, femaleNames));
        }

        sb.append("=".repeat(80) + "\n");
        sb.append("排班说明：\n");
        sb.append("1. 周一至周五正常排班，节假日不排班\n");
        sb.append("2. 每天男生4人，女生2人\n");
        sb.append("3. 优先安排纪检委员\n");
        sb.append("4. 每人一周上班次数上限为5次\n");

        return sb.toString();
    }

    /**
     * 生成2026年中国法定节假日
     */
    public Set<LocalDate> generateHolidays2026() {
        Set<LocalDate> holidays = new LinkedHashSet<>();

        // 元旦
        holidays.add(LocalDate.of(2026, 1, 1));
        holidays.add(LocalDate.of(2026, 1, 2));
        holidays.add(LocalDate.of(2026, 1, 3));

        // 春节（假设）
        holidays.add(LocalDate.of(2026, 2, 14));
        holidays.add(LocalDate.of(2026, 2, 15));
        holidays.add(LocalDate.of(2026, 2, 16));
        holidays.add(LocalDate.of(2026, 2, 17));
        holidays.add(LocalDate.of(2026, 2, 18));
        holidays.add(LocalDate.of(2026, 2, 19));
        holidays.add(LocalDate.of(2026, 2, 20));

        // 清明节
        holidays.add(LocalDate.of(2026, 4, 4));
        holidays.add(LocalDate.of(2026, 4, 5));
        holidays.add(LocalDate.of(2026, 4, 6));

        // 劳动节
        holidays.add(LocalDate.of(2026, 5, 1));
        holidays.add(LocalDate.of(2026, 5, 2));
        holidays.add(LocalDate.of(2026, 5, 3));
        holidays.add(LocalDate.of(2026, 5, 4));
        holidays.add(LocalDate.of(2026, 5, 5));

        // 端午节
        holidays.add(LocalDate.of(2026, 6, 19));
        holidays.add(LocalDate.of(2026, 6, 20));
        holidays.add(LocalDate.of(2026, 6, 21));

        // 中秋节
        holidays.add(LocalDate.of(2026, 9, 25));
        holidays.add(LocalDate.of(2026, 9, 26));
        holidays.add(LocalDate.of(2026, 9, 27));

        // 国庆节
        holidays.add(LocalDate.of(2026, 10, 1));
        holidays.add(LocalDate.of(2026, 10, 2));
        holidays.add(LocalDate.of(2026, 10, 3));
        holidays.add(LocalDate.of(2026, 10, 4));
        holidays.add(LocalDate.of(2026, 10, 5));
        holidays.add(LocalDate.of(2026, 10, 6));
        holidays.add(LocalDate.of(2026, 10, 7));
        holidays.add(LocalDate.of(2026, 10, 8));

        return holidays;
    }

    /**
     * 生成一周的排班表（周一到周五）
     * 用于简单排班场景
     */
    public List<DutySchedule> generateWeeklySchedule(List<Person> persons) {
        logger.info("开始生成一周排班表");
        
        List<DutySchedule> schedules = new ArrayList<>();
        String[] dayNames = {"周一", "周二", "周三", "周四", "周五"};
        
        // 按性别分组
        Map<Boolean, List<Person>> maleGroups = persons.stream()
                .filter(p -> "男".equals(p.getGender()))
                .collect(Collectors.groupingBy(Person::isInspector));

        Map<Boolean, List<Person>> femaleGroups = persons.stream()
                .filter(p -> "女".equals(p.getGender()))
                .collect(Collectors.groupingBy(Person::isInspector));

        List<Person> maleInspectors = maleGroups.getOrDefault(true, new ArrayList<>());
        List<Person> maleNormal = maleGroups.getOrDefault(false, new ArrayList<>());
        List<Person> femaleInspectors = femaleGroups.getOrDefault(true, new ArrayList<>());
        List<Person> femaleNormal = femaleGroups.getOrDefault(false, new ArrayList<>());

        logger.info("人员统计 - 男生: {}, 女生: {}", 
                maleNormal.size() + maleInspectors.size(),
                femaleNormal.size() + femaleInspectors.size());

        // 生成周一到周五的排班
        for (int dayIdx = 0; dayIdx < 5; dayIdx++) {
            List<Person> maleDuty = selectPersons(maleInspectors, maleNormal, MALE_PER_DAY, dayIdx);
            List<Person> femaleDuty = selectPersons(femaleInspectors, femaleNormal, FEMALE_PER_DAY, dayIdx);

            DutySchedule schedule = new DutySchedule(
                    java.sql.Date.valueOf(LocalDate.now().with(TemporalAdjusters.previousOrSame(DayOfWeek.MONDAY)).plusDays(dayIdx)),
                    false,
                    maleDuty,
                    femaleDuty
            );
            schedules.add(schedule);

            logger.info("{}: 男生={}, 女生={}", dayNames[dayIdx],
                    maleDuty.stream().map(Person::getName).collect(Collectors.joining(",")),
                    femaleDuty.stream().map(Person::getName).collect(Collectors.joining(",")));
        }

        logger.info("一周排班生成完成");
        return schedules;
    }

    /**
     * 将排班结果导出为Excel工作簿（格式：每周一行，每列显示具体学生姓名）
     * 列格式：周次 | 男1 | 男2 | 男3 | 男4 | 女1 | 女2
     */
    public byte[] createWeeklyScheduleWorkbook(List<DutySchedule> schedules) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("排班表");

            // Header: 周次 | 男1 | 男2 | 男3 | 男4 | 女1 | 女2
            int rowIdx = 0;
            Row header = sheet.createRow(rowIdx++);
            header.createCell(0).setCellValue("周次");
            header.createCell(1).setCellValue("男1");
            header.createCell(2).setCellValue("男2");
            header.createCell(3).setCellValue("男3");
            header.createCell(4).setCellValue("男4");
            header.createCell(5).setCellValue("女1");
            header.createCell(6).setCellValue("女2");

            String[] dayLabels = {"周一", "周二", "周三", "周四", "周五"};

            for (int i = 0; i < schedules.size(); i++) {
                DutySchedule s = schedules.get(i);
                Row r = sheet.createRow(rowIdx++);
                r.createCell(0).setCellValue(dayLabels[i]);

                // 填充男生姓名（4列）
                List<Person> males = s.getMaleDuty();
                for (int j = 0; j < 4; j++) {
                    String name = (j < males.size() && males.get(j) != null) ? males.get(j).getName() : "";
                    r.createCell(1 + j).setCellValue(name);
                }

                // 填充女生姓名（2列）
                List<Person> females = s.getFemaleDuty();
                for (int j = 0; j < 2; j++) {
                    String name = (j < females.size() && females.get(j) != null) ? females.get(j).getName() : "";
                    r.createCell(5 + j).setCellValue(name);
                }
            }

            // 自动调整列宽
            sheet.setColumnWidth(0, 12 * 256);
            for (int i = 1; i <= 6; i++) {
                sheet.setColumnWidth(i, 15 * 256);
            }

            workbook.write(out);
            return out.toByteArray();
        }
    }

    /**
     * 将排班表写入一个 Excel 工作簿，并返回字节数组
     */
    public byte[] createScheduleWorkbook(List<DutySchedule> schedules) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("排班表");

            // Header: 日期 | 是否节假日 | 男1 | 男2 | 男3 | 男4 | 女1 | 女2
            int rowIdx = 0;
            Row header = sheet.createRow(rowIdx++);
            header.createCell(0).setCellValue("日期");
            header.createCell(1).setCellValue("是否节假日");
            header.createCell(2).setCellValue("男1");
            header.createCell(3).setCellValue("男2");
            header.createCell(4).setCellValue("男3");
            header.createCell(5).setCellValue("男4");
            header.createCell(6).setCellValue("女1");
            header.createCell(7).setCellValue("女2");

            for (DutySchedule s : schedules) {
                Row r = sheet.createRow(rowIdx++);
                r.createCell(0).setCellValue(s.getDate().toString());
                r.createCell(1).setCellValue(s.isHoliday() ? "是" : "否");

                if (s.isHoliday()) {
                    for (int c = 2; c <= 7; c++) r.createCell(c).setCellValue("-");
                } else {
                    // 填充男生 4 列
                    List<Person> males = s.getMaleDuty();
                    for (int i = 0; i < 4; i++) {
                        String name = (i < males.size() && males.get(i) != null) ? males.get(i).getName() : "";
                        r.createCell(2 + i).setCellValue(name);
                    }

                    // 填充女生 2 列
                    List<Person> females = s.getFemaleDuty();
                    for (int i = 0; i < 2; i++) {
                        String name = (i < females.size() && females.get(i) != null) ? females.get(i).getName() : "";
                        r.createCell(6 + i).setCellValue(name);
                    }
                }
            }

            // 自动调整列宽
            for (int i = 0; i <= 7; i++) sheet.autoSizeColumn(i);

            workbook.write(out);
            return out.toByteArray();
        }
    }
}
