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

/**
 * 排班服务
 */
@Service
public class ScheduleService {
    private static final Logger logger = LoggerFactory.getLogger(ScheduleService.class);
    private static final int MALE_PER_DAY = 4;
    private static final int FEMALE_PER_DAY = 2;

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
                List<Person> maleDuty = selectPersons(maleInspectors, maleNormal, MALE_PER_DAY);
                List<Person> femaleDuty = selectPersons(femaleInspectors, femaleNormal, FEMALE_PER_DAY);

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
    private List<Person> selectPersons(List<Person> inspectors, List<Person> normal, int count) {
        List<Person> selected = new ArrayList<>();
        List<Person> all = new ArrayList<>();

        // 优先选择纪检委员（按已排班次数升序排序）
        inspectors.sort(Comparator.comparingInt(Person::getWeeklyCount)
                .thenComparingInt(Person::getMonthlyCount)
                .reversed());
        all.addAll(inspectors);

        // 添加普通人员
        normal.sort(Comparator.comparingInt(Person::getWeeklyCount)
                .thenComparingInt(Person::getMonthlyCount)
                .reversed());
        all.addAll(normal);

        // 选择符合条件的人员
        for (Person person : all) {
            if (selected.size() >= count) break;

            // 检查约束条件
            if (person.getWeeklyCount() < 2 && person.getMonthlyCount() < 4) {
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
        sb.append("4. 每人一周不超过2次，一个月不超过4次\n");

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
}
