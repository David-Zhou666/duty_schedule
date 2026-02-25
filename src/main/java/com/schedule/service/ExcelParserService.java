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
     * 支持格式1（有性别列）：
     * 第一列：姓名，第二列：性别（男/女），第三列：是否纪检委员，第四-八列：周一-周五可用性
     * 支持格式2（无性别列，由Sheet名判断）：
     * 第一列：姓名，第二-六列：周一-周五可用性
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

                    Person person = parseRow(row, sheetName);
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
    private Person parseRow(Row row, String sheetName) {
        try {
            Cell nameCell = row.getCell(0);

            if (nameCell == null || nameCell.getStringCellValue() == null ||
                nameCell.getStringCellValue().trim().isEmpty()) {
                return null;
            }

            String name = nameCell.getStringCellValue().trim();
            
            // 判断是否有性别列（第二列有值且是男/女）
            Cell secondCell = row.getCell(1);
            String secondValue = getCellValue(secondCell);
            boolean hasGenderColumn = "男".equals(secondValue) || "女".equals(secondValue);
            
            String gender;
            boolean isInspector;
            int availabilityStartCol;
            
            if (hasGenderColumn) {
                // 格式1：有性别列
                gender = secondValue;
                String inspectorStr = getCellValue(row.getCell(2));
                isInspector = "是".equals(inspectorStr) || "Y".equalsIgnoreCase(inspectorStr) || "true".equalsIgnoreCase(inspectorStr);
                availabilityStartCol = 3;
            } else {
                // 格式2：无性别列，根据sheet名判断
                gender = sheetName.contains("男") || sheetName.contains("boy") ? "男" : "女";
                isInspector = false; // 默认不是纪检委员
                availabilityStartCol = 1;
            }

            Person person = new Person(name, gender, isInspector);

            // 读取周一到周五可用性（5列）
            boolean[] avail = new boolean[5];
            for (int i = 0; i < 5; i++) {
                Cell c = row.getCell(availabilityStartCol + i);
                String v = getCellValue(c);
                boolean ok = false;
                if (v != null && !v.isEmpty()) {
                    if ("是".equals(v) || "Y".equalsIgnoreCase(v) || "true".equalsIgnoreCase(v) || "1".equals(v)) {
                        ok = true;
                    }
                } else {
                    // 空值视为可用
                    ok = true;
                }
                avail[i] = ok;
            }
            person.setAvailability(avail);

            return person;

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
