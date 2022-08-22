import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.Scanner;

public class Main {
    static Scanner scanner = new Scanner(System.in);

    public static void main(String[] args) {
        System.out.print("Enter the path: ");
        parser(scanner.next());
    }

    public static void parser(String path) {
        try (BufferedReader reader = new BufferedReader(new FileReader(path, StandardCharsets.UTF_8)); XSSFWorkbook workbook = new XSSFWorkbook()) {
            int rowIndex = 0;
            String line = reader.readLine();
            String currentSheetName = ".";
            XSSFSheet sheet = null;
            while (line != null) {
                String[] cellData = line.replaceAll("\\s+", " ").split(" ");
                if (cellData.length == 8) {
                    assert sheet != null;
                    setCellValue(sheet, rowIndex++, cellData);
                } else if (cellData.length == 7 && !currentSheetName.equals(cellData[6] = cellData[6].substring(0, cellData[6].indexOf('/')))) {
                    sheet = workbook.createSheet(cellData[6]);
                    currentSheetName = cellData[6];
                    rowIndex = 1;
                    createSheet(sheet);
                }
                line = reader.readLine();
            }
            workbook.write(new FileOutputStream("output.xlsx"));
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

    public static void setCellValue(Sheet sheet, int rowIndex, String[] cellData) {
        int toDate = 4;
        Row currentRow = sheet.createRow(rowIndex);
        sheet.autoSizeColumn(rowIndex - 1);
        int i = 0;
        int cellIndex = 1;
        for (; cellIndex < toDate; ++i, ++cellIndex) {
            currentRow.createCell(i).setCellValue(cellData[cellIndex]);
        }
        currentRow.createCell(i++).setCellValue(cellData[cellIndex] + ' ' + cellData[++cellIndex] + ' ' + cellData[++cellIndex]);
        currentRow.createCell(i).setCellValue("http://virtual2057/svn/mc21-dcf-dev/tag/UAC240000312A02/" + cellData[++cellIndex]);
    }

    public static void createSheet(Sheet sheet) {
        Row currentRow = sheet.createRow(0);
        sheet.autoSizeColumn(0);
        currentRow.createCell(0).setCellValue("№");
        currentRow.createCell(1).setCellValue("Author");
        currentRow.createCell(2).setCellValue("Size");
        currentRow.createCell(3).setCellValue("Date");
        currentRow.createCell(4).setCellValue("Path");
    }
}
