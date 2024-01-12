package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelAnalyzer {

    public static void main(String[] args) {
        try {
            String filePath = "src\\Assignment_Timecard.xlsx";
            analyzeExcelFile(filePath, 7);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void analyzeExcelFile(String filePath, int consecutiveDaysThreshold) throws IOException {
        FileInputStream file = new FileInputStream(new File(filePath));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);

        Set<String> consecutivePrinted = new HashSet<>();
        Set<String> shortBreakPrinted = new HashSet<>();
        Set<String> longShiftPrinted = new HashSet<>();

        System.out.println("part 1: ");
        checkConsecutiveDays(sheet, consecutiveDaysThreshold, consecutivePrinted);

        System.out.println("part 2: ");
        checkShortBreaks(sheet, shortBreakPrinted);

        System.out.println("part 3: ");
        checkLongShifts(sheet, longShiftPrinted);

        workbook.close();
        file.close();
    }

    private static void checkConsecutiveDays(Sheet sheet, int consecutiveDaysThreshold, Set<String> printed) {
        for (int index = 1; index < sheet.getLastRowNum(); index++) {
            Row currentRow = sheet.getRow(index);
            String employeeName = getStringValue(currentRow.getCell(7)); // Assuming 'Employee Name' is in the 8th column
            String positionId = getStringValue(currentRow.getCell(0));  // Assuming 'Position ID' is in the 1st column

            if (printed.contains(employeeName)) {
                continue;
            }

            if (index > 0 && employeeName.equals(getStringValue(sheet.getRow(index - 1).getCell(7)))) {
                int consecutiveDays = 1;
                for (int i = index - 1; i >= 0; i--) {
                    if (employeeName.equals(getStringValue(sheet.getRow(i).getCell(7)))) {
                        consecutiveDays++;
                    } else {
                        break;
                    }
                }
                if (consecutiveDays >= consecutiveDaysThreshold) {
                    System.out.println("Employee: " + employeeName + ", Position: " + positionId);
                    printed.add(employeeName);
                }
            }
        }
    }

    private static void checkShortBreaks(Sheet sheet, Set<String> printed) {
        Map<String, Date> employeeBreaks = new HashMap<>();

        for (int index = 1; index <= sheet.getLastRowNum(); index++) {
            Row currentRow = sheet.getRow(index);
            String employeeName = getStringValue(currentRow.getCell(7)); // Assuming 'Employee Name' is in the 8th column
            String positionId = getStringValue(currentRow.getCell(0));  // Assuming 'Position ID' is in the 1st column

            if (printed.contains(employeeName)) {
                continue;
            }

            if (employeeBreaks.containsKey(employeeName)) {
                Date lastTimeOut = employeeBreaks.get(employeeName);
                Date timeIn = getDateValue(currentRow.getCell(2), employeeBreaks); // Assuming 'Time' is in the 3rd column

                if (timeIn != null && lastTimeOut != null) {
                    long timeDiff = (timeIn.getTime() - lastTimeOut.getTime()) / (60 * 60 * 1000);
                    if (1 < timeDiff && timeDiff < 10) {
                        System.out.println("Employee: " + employeeName + ", Position: " + positionId);
                        printed.add(employeeName);
                    }
                }
            }

            employeeBreaks.put(employeeName, getDateValue(currentRow.getCell(3), employeeBreaks)); // Assuming 'Time Out' is in the 4th column
        }
    }

    private static void checkLongShifts(Sheet sheet, Set<String> printed) {
        for (int index = 0; index <= sheet.getLastRowNum(); index++) {
            Row currentRow = sheet.getRow(index);
            String employeeName = getStringValue(currentRow.getCell(7)); // Assuming 'Employee Name' is in the 8th column
            String positionId = getStringValue(currentRow.getCell(0));  // Assuming 'Position ID' is in the 1st column

            if (printed.contains(employeeName)) {
                continue;
            }

            String durationStr = getStringValue(currentRow.getCell(4)); // Assuming 'Timecard Hours (as Time)' is in the 5th column
            if (durationStr != null) {
                try {
                    String[] parts = durationStr.split(":");
                    int hours = Integer.parseInt(parts[0]);
                    int minutes = Integer.parseInt(parts[1]);
                    int totalMinutes = hours * 60 + minutes;
                    if (totalMinutes > 840) { // 14 hours in minutes
                        System.out.println("Employee: " + employeeName + ", Position: " + positionId);
                        printed.add(employeeName);
                    }
                } catch (NumberFormatException | ArrayIndexOutOfBoundsException e) {
                    // Handle invalid duration format
                }
            }
        }
    }

    private static String getStringValue(Cell cell) {
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    return String.valueOf(cell.getNumericCellValue());
                default:
                    // Handle other cell types if needed
                    return null;
            }
        }
        return null;
    }
    
    private static Date getDateValue(Cell cell, Map<String, Date> dateCache) {
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    String dateString = cell.getStringCellValue().trim(); // Trim to handle leading/trailing whitespaces
                    if (!dateString.isEmpty()) {
                        try {
                            return new SimpleDateFormat("MM/dd/yyyy HH:mm a").parse(dateString);
                        } catch (ParseException e) {
                            e.printStackTrace(); // Handle parsing exception as needed
                        }
                    }
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue();
                    }
                    break;
                // Handle other cell types if needed
            }
        }
        return null;
    }
    
    

}
