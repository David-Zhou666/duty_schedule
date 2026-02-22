package com.schedule.service;

import com.schedule.model.Person;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.*;

/**
 * Excel 文件解析服务
 */
@Service
public class ExcelParserService {
    private static final Logger logger = LoggerFactory.getLogger(ExcelParserService.class);

    /**
     * 解析 Excel 文件，支持多 Sheet 页
     * 假设 Excel 格式：
     * 第一列：姓名
     * 第二列：性别（男/女）
     * 第三列：是否纪检委员（是/否）
     */
    public List<Person> parseExcelFile(MultipartFile file) throws IOException {
        List<Person> persons = new ArrayList<>();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            // 遍历所有 Sheet
            int sheetCount = workbook.getNumberOfSheets();
            logger.info("Excel 文件共有 {} 个 Sheet 页", sheetCount);

            for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                String sheetName = sheet.getSheetName();
                logger.info("正在解析 Sheet: {}", sheetName);

                // 从第二行开始读取（假设第一行是表头）
                for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row == null) continue;

                    Person person = parseRow(row);
                    if (person != null) {
                        persons.add(person);
                        logger.debug("解析到人员: {}", person);
                    }
                }
            }
        }

        logger.info("总共解析到 {} 名人员", persons.size());
        return persons;
    }

    /**
     * 解析单行数据
     */
    private Person parseRow(Row row) {
        try {
            Cell nameCell = row.getCell(0);
            Cell genderCell = row.getCell(1);
            Cell inspectorCell = row.getCell(2);

            if (nameCell == null || nameCell.getStringCellValue() == null ||
                nameCell.getStringCellValue().trim().isEmpty()) {
                return null;
            }

            String name = nameCell.getStringCellValue().trim();
            String gender = getCellValue(genderCell);
            String inspector = getCellValue(inspectorCell);

            // 验证性别
            if (!"男".equals(gender) && !"女".equals(gender)) {
                logger.warn("人员 {} 的性别值异常: {}，默认设为男", name, gender);
                gender = "男";
            }

            // 验证是否纪检委员
            boolean isInspector = "是".equals(inspector) || "Y".equalsIgnoreCase(inspector) || "true".equalsIgnoreCase(inspector);

            return new Person(name, gender, isInspector);

        } catch (Exception e) {
            logger.error("解析行 {} 时出错: {}", row.getRowNum(), e.getMessage());
            return null;
        }
    }

    /**
     * 获取单元格字符串值
     */
    private String getCellValue(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
