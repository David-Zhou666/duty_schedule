import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 独立排班程序
 * 读取 duty_temple.xlsx 文件，生成排班表
 */
public class GenerateSchedule {

    static class Person {
        String name;
        String gender;
        boolean[] availability = new boolean[5]; // 周一到周五
        int weeklyCount = 0;

        Person(String name, String gender) {
            this.name = name;
            this.gender = gender;
            Arrays.fill(availability, true); // 默认可用
        }

        boolean isAvailable(int day) { // 0=周一, 4=周五
            return availability[day];
        }
    }

    public static void main(String[] args) {
        String filePath = "src/main/resources/duty_temple.xlsx";
        List<Person> persons = readExcel(filePath);

        if (persons.isEmpty()) {
            System.out.println("未读取到任何人员信息");
            return;
        }

        System.out.println("成功读取 " + persons.size() + " 名人员");

        // 分离男生和女生
        List<Person> males = new ArrayList<>();
        List<Person> females = new ArrayList<>();
        for (Person p : persons) {
            if ("男".equals(p.gender)) males.add(p);
            else females.add(p);
        }

        System.out.println("男生: " + males.size() + " 人");
        System.out.println("女生: " + females.size() + " 人");

        // 生成一周排班表
        String[] dayNames = {"周一", "周二", "周三", "周四", "周五"};
        List<List<Person>> maleSchedules = new ArrayList<>();
        List<List<Person>> femaleSchedules = new ArrayList<>();

        for (int day = 0; day < 5; day++) {
            List<Person> maleDuty = selectDuty(males, 4, day);
            List<Person> femaleDuty = selectDuty(females, 2, day);
            maleSchedules.add(maleDuty);
            femaleSchedules.add(femaleDuty);
        }

        // 输出排班表
        System.out.println("\n========================================");
        System.out.println("            值日排班表");
        System.out.println("========================================");
        System.out.printf("%-8s %-12s %-12s %-12s %-12s %-12s %-12s\n",
                "周次", "男1", "男2", "男3", "男4", "女1", "女2");
        System.out.println("----------------------------------------");

        for (int day = 0; day < 5; day++) {
            System.out.printf("%-8s", dayNames[day]);
            for (Person p : maleSchedules.get(day)) {
                System.out.printf(" %-12s", p != null ? p.name : "");
            }
            for (Person p : femaleSchedules.get(day)) {
                System.out.printf(" %-12s", p != null ? p.name : "");
            }
            System.out.println();
        }
        System.out.println("========================================");

        // 统计每人本周值日次数
        System.out.println("\n男生本周值日次数:");
        for (Person p : males) {
            if (p.weeklyCount > 0) {
                System.out.println("  " + p.name + ": " + p.weeklyCount + " 次");
            }
        }
        System.out.println("\n女生本周值日次数:");
        for (Person p : females) {
            if (p.weeklyCount > 0) {
                System.out.println("  " + p.name + ": " + p.weeklyCount + " 次");
            }
        }

        // 导出Excel文件
        exportToExcel(dayNames, maleSchedules, femaleSchedules);
        System.out.println("\n排班表已导出到: duty_schedule_result.xlsx");
    }

    /**
     * 从Excel读取人员信息
     */
    static List<Person> readExcel(String filePath) {
        List<Person> persons = new ArrayList<>();

        File file = new File(filePath);
        if (!file.exists()) {
            System.err.println("文件不存在: " + filePath);
            return persons;
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            int sheetCount = workbook.getNumberOfSheets();
            System.out.println("Excel文件共有 " + sheetCount + " 个Sheet页");

            for (int sheetIdx = 0; sheetIdx < sheetCount; sheetIdx++) {
                Sheet sheet = workbook.getSheetAt(sheetIdx);
                String sheetName = sheet.getSheetName();
                System.out.println("正在读取 Sheet: " + sheetName);

                // 根据sheet名判断性别
                boolean isMale = sheetName.contains("男");

                // 从第二行开始（第一行是表头）
                for (int rowIdx = 1; rowIdx <= sheet.getLastRowNum(); rowIdx++) {
                    Row row = sheet.getRow(rowIdx);
                    if (row == null) continue;

                    Cell nameCell = row.getCell(0);
                    String name = getCellValue(nameCell);

                    if (name == null || name.trim().isEmpty()) continue;

                    name = name.trim();
                    String gender = isMale ? "男" : "女";

                    Person person = new Person(name, gender);

                    // 读取周一到周五的可用性（列1-5）
                    for (int dayIdx = 0; dayIdx < 5; dayIdx++) {
                        Cell cell = row.getCell(dayIdx + 1);
                        String value = getCellValue(cell);

                        // "是"或其他值视为可用，空值也视为可用
                        boolean available = value == null || value.isEmpty() || "是".equals(value);
                        person.availability[dayIdx] = available;
                    }

                    persons.add(person);
                    System.out.println("  " + name + " (" + gender + ")");
                }
            }

        } catch (IOException e) {
            System.err.println("读取Excel文件失败: " + e.getMessage());
            e.printStackTrace();
        }

        return persons;
    }

    /**
     * 选择值班人员
     */
    static List<Person> selectDuty(List<Person> pool, int count, int dayOfWeek) {
        List<Person> selected = new ArrayList<>();

        // 筛选当天可用的人
        List<Person> available = new ArrayList<>();
        for (Person p : pool) {
            if (p.isAvailable(dayOfWeek)) {
                available.add(p);
            }
        }

        // 按已排班次数排序
        available.sort((a, b) -> a.weeklyCount - b.weeklyCount);

        // 选择符合条件的人（每周不超过5次）
        for (Person p : available) {
            if (selected.size() >= count) break;
            if (p.weeklyCount < 5) {
                selected.add(p);
                p.weeklyCount++;
            }
        }

        // 如果人数不足，放宽约束
        if (selected.size() < count) {
            for (Person p : available) {
                if (selected.size() >= count) break;
                if (!selected.contains(p)) {
                    selected.add(p);
                    p.weeklyCount++;
                }
            }
        }

        return selected;
    }

    /**
     * 获取单元格值
     */
    static String getCellValue(Cell cell) {
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
     * 导出Excel文件
     */
    static void exportToExcel(String[] dayNames, List<List<Person>> maleSchedules,
                              List<List<Person>> femaleSchedules) {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream out = new FileOutputStream("duty_schedule_result.xlsx")) {

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

            for (int day = 0; day < 5; day++) {
                Row r = sheet.createRow(rowIdx++);
                r.createCell(0).setCellValue(dayNames[day]);

                // 填充男生姓名
                List<Person> males = maleSchedules.get(day);
                for (int j = 0; j < 4; j++) {
                    String name = (j < males.size() && males.get(j) != null) ? males.get(j).name : "";
                    r.createCell(1 + j).setCellValue(name);
                }

                // 填充女生姓名
                List<Person> females = femaleSchedules.get(day);
                for (int j = 0; j < 2; j++) {
                    String name = (j < females.size() && females.get(j) != null) ? females.get(j).name : "";
                    r.createCell(5 + j).setCellValue(name);
                }
            }

            // 自动调整列宽
            sheet.setColumnWidth(0, 12 * 256);
            for (int i = 1; i <= 6; i++) {
                sheet.setColumnWidth(i, 15 * 256);
            }

            workbook.write(out);

        } catch (IOException e) {
            System.err.println("导出Excel失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
