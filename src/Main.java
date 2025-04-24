// Java package imports
import java.util.Scanner;
import java.util.HashMap;
import java.util.TreeMap;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Pattern;

// Apache POI imports for making spreadsheets
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.awt.Color;

public class Main {

    static HashMap<Integer, Scanner> day1 = new HashMap<>();
    static TreeMap<String, String[]> day1info = new TreeMap<>();
    static TreeMap<String, HashMap<Integer, Boolean>> day1attendance = new TreeMap<>();

    static HashMap<Integer, Scanner> day2 = new HashMap<>();
    static TreeMap<String, String[]> day2info = new TreeMap<>();
    static TreeMap<String, HashMap<Integer, Boolean>> day2attendance = new TreeMap<>();

    static String regex = ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)";

    public static void main(String[] args) throws FileNotFoundException {
        // path where all csv files of attendance is stored, just change it if needed.
        String csvPath = "C:/Users/kurut0/Desktop/Seminar Attendance/";

        // day 1 seminar
        System.out.println("Day 1 Tally:");
        openDay(day1, csvPath + "Day 1/", 8);
        processDay(day1, day1attendance, day1info, 8, csvPath + "Day1_Attendance.xlsx");

        // day 2 seminar
        System.out.println("Day 2 Tally:");
        openDay(day2, csvPath + "Day 2/", 6);
        processDay(day2, day2attendance, day2info, 6, csvPath + "Day2_Attendance.xlsx");
    }

    static void openDay(HashMap<Integer, Scanner> day, String path, int numSeminars) throws FileNotFoundException {
        for(int i = 1; i <= numSeminars; i++) {
            day.put(i, new Scanner(new FileReader(path + i + ".csv")));
        }
    }

    static void processDay(HashMap<Integer, Scanner> dayFile, TreeMap<String, HashMap<Integer, Boolean>> attendance, TreeMap<String, String[]> info, int numSeminars, String path) {
        Workbook workbook = new XSSFWorkbook();

        Sheet bsitSheet = workbook.createSheet("BSIT"); int bsitRow = 1; createHeaders(workbook, bsitSheet, numSeminars);
        Sheet bscsSheet = workbook.createSheet("BSCS"); int bscsRow = 1; createHeaders(workbook, bscsSheet, numSeminars);
        Sheet bsisSheet = workbook.createSheet("BSIS"); int bsisRow = 1; createHeaders(workbook, bsisSheet, numSeminars);
        Sheet bsemcSheet = workbook.createSheet("BSEMC"); int bsemcRow = 1; createHeaders(workbook, bsemcSheet, numSeminars);
        Sheet miscSheet = workbook.createSheet("Misc"); int miscRow = 1; createHeaders(workbook, miscSheet, numSeminars);

        CellStyle presentStyle = workbook.createCellStyle();
        XSSFColor presentColor = new XSSFColor(Color.decode("#77DD77"), null);
        presentStyle.setFillForegroundColor(presentColor);
        presentStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle absentStyle = workbook.createCellStyle();
        XSSFColor absentColor = new XSSFColor(Color.decode("#FF6961"), null);
        absentStyle.setFillForegroundColor(absentColor);
        absentStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        int attendanceCol = 0;

        for(int i = 1; i <= numSeminars; i++) { // loop for each seminar
            Scanner data = dayFile.get(i);
            data.nextLine();

            while (data.hasNext()) { // loop for each student
                String rowData = data.nextLine();

                Pattern splitPattern = Pattern.compile(regex);
                String[] attendanceArray = splitPattern.split(rowData, -1);

                // Index Reference: 1 -> email, 2 -> full name, 3 -> section
                if (!info.containsKey(attendanceArray[1]))
                    info.put(attendanceArray[1], new String[] {attendanceArray[2].toUpperCase().replaceAll("^\"|\"$", ""), attendanceArray[3].toUpperCase()});

                HashMap<Integer, Boolean> currentAttended = new HashMap<>();
                if (!attendance.containsKey(attendanceArray[1])) {
                    for(int val = 1; val <= numSeminars; val++) {
                        currentAttended.put(val, null);
                    }
                    currentAttended.put(i, true);
                    attendance.put(attendanceArray[1], currentAttended);
                }
                else {
                    currentAttended = attendance.get(attendanceArray[1]);
                    currentAttended.put(i, true);
                    attendance.put(attendanceArray[1], currentAttended);
                }
            }
        }

        int currentStudent = 0;
        for(HashMap.Entry<String, String[]> entry : info.entrySet()) {

            if(entry.getValue()[1].contains("BSIT")) {
                Row row = bsitSheet.createRow(bsitRow++);

                row.createCell(0).setCellValue(entry.getKey());

                for (int val = 0; val < entry.getValue().length; val++) {
                    row.createCell(val + 1).setCellValue(entry.getValue()[val]);
                    attendanceCol = val+2;
                }

                for (int val = 0; val < attendance.get(entry.getKey()).size(); val++) {
                    if(attendance.get(entry.getKey()).get(val+1) != null) {
                        Cell cell = row.createCell(val + attendanceCol);
                        cell.setCellValue("Present");
                        cell.setCellStyle(presentStyle);
                    }
                    else {
                        Cell cell = row.createCell(val + attendanceCol);
                        cell.setCellValue("Absent");
                        cell.setCellStyle(absentStyle);
                    }
                }
            }
            else if(entry.getValue()[1].contains("BSCS")) {
                Row row = bscsSheet.createRow(bscsRow++);

                row.createCell(0).setCellValue(entry.getKey());

                for (int val = 0; val < entry.getValue().length; val++) {
                    row.createCell(val + 1).setCellValue(entry.getValue()[val]);
                    attendanceCol = val+2;
                }

                for (int val = 0; val < attendance.get(entry.getKey()).size(); val++) {
                    if(attendance.get(entry.getKey()).get(val+1) != null) {
                        Cell cell = row.createCell(val + attendanceCol);
                        cell.setCellValue("Present");
                        cell.setCellStyle(presentStyle);
                    }
                    else {
                        Cell cell = row.createCell(val + attendanceCol);
                        cell.setCellValue("Absent");
                        cell.setCellStyle(absentStyle);
                    }
                }
            }
            else if(entry.getValue()[1].contains("BSIS")) {
                Row row = bsisSheet.createRow(bsisRow++);

                row.createCell(0).setCellValue(entry.getKey());

                for (int val = 0; val < entry.getValue().length; val++) {
                    row.createCell(val + 1).setCellValue(entry.getValue()[val]);
                    attendanceCol = val+2;
                }

                for (int val = 0; val < attendance.get(entry.getKey()).size(); val++) {
                    if(attendance.get(entry.getKey()).get(val+1) != null) {
                        Cell cell = row.createCell(val + attendanceCol);
                        cell.setCellValue("Present");
                        cell.setCellStyle(presentStyle);
                    }
                    else {
                        Cell cell = row.createCell(val + attendanceCol);
                        cell.setCellValue("Absent");
                        cell.setCellStyle(absentStyle);
                    }
                }
            }
            else if(entry.getValue()[1].contains("BSEMC")) {
                Row row = bsemcSheet.createRow(bsemcRow++);

                row.createCell(0).setCellValue(entry.getKey());

                for (int val = 0; val < entry.getValue().length; val++) {
                    row.createCell(val + 1).setCellValue(entry.getValue()[val]);
                    attendanceCol = val+2;
                }

                for (int val = 0; val < attendance.get(entry.getKey()).size(); val++) {
                    if(attendance.get(entry.getKey()).get(val+1) != null) {
                        Cell cell = row.createCell(val + attendanceCol);
                        cell.setCellValue("Present");
                        cell.setCellStyle(presentStyle);
                    }
                    else {
                        Cell cell = row.createCell(val + attendanceCol);
                        cell.setCellValue("Absent");
                        cell.setCellStyle(absentStyle);
                    }
                }
            }
            else {
                Row row = miscSheet.createRow(miscRow++);

                row.createCell(0).setCellValue(entry.getKey());

                for (int val = 0; val < entry.getValue().length; val++) {
                    row.createCell(val + 1).setCellValue(entry.getValue()[val]);
                    attendanceCol = val+2;
                }

                for (int val = 0; val < attendance.get(entry.getKey()).size(); val++) {
                    if(attendance.get(entry.getKey()).get(val+1) != null) {
                        Cell cell = row.createCell(val + attendanceCol);
                        cell.setCellValue("Present");
                        cell.setCellStyle(presentStyle);
                    }
                    else {
                        Cell cell = row.createCell(val + attendanceCol);
                        cell.setCellValue("Absent");
                        cell.setCellStyle(absentStyle);
                    }
                }
            }

            currentStudent++;
        }

        for(int i = 0; i < (numSeminars + 3); i++) {
            bsitSheet.autoSizeColumn(i);
            bscsSheet.autoSizeColumn(i);
            bsisSheet.autoSizeColumn(i);
            bsemcSheet.autoSizeColumn(i);
            miscSheet.autoSizeColumn(i);
        }

        System.out.println("Total number of attendees: " +  info.size() + "\n"); // gets total students of the day based on student info

        try (FileOutputStream fileOut = new FileOutputStream(path)) {
            workbook.write(fileOut);
            workbook.close();
            System.out.println("Workbook created successfully");
        } catch(IOException e) {
            System.out.println("ERROR: Workbook creation failed. Please see error stack trace below:");
            e.printStackTrace();
        }
    }

    public static void createHeaders(Workbook workbook, Sheet worksheet, int numSeminars) {
        CellStyle headerStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short)14);
        headerStyle.setFont(font);

        Row headerRow = worksheet.createRow(0);
        Cell emailCell = headerRow.createCell(0);
        emailCell.setCellValue("Email Address");
        emailCell.setCellStyle(headerStyle);

        Cell nameCell = headerRow.createCell(1);
        nameCell.setCellValue("Full Name");
        nameCell.setCellStyle(headerStyle);

        Cell sectionCell = headerRow.createCell(2);
        sectionCell.setCellValue("Section");
        sectionCell.setCellStyle(headerStyle);

        for(int i = 1; i <= numSeminars; i++) {
            Cell topicCell = headerRow.createCell(2+i);
            topicCell.setCellValue("Topic " + i);
            topicCell.setCellStyle(headerStyle);
        }
    }
}