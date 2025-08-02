package service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.nio.file.*;
import java.util.ArrayList;
import java.util.List;

@Service
public class DatToExcelService {

    public void convertDatToExcel(String inputPath, String inputFileName, String outputPath) throws IOException {
        Path inputFile = Paths.get(inputPath, inputFileName);

        String outputFileName = inputFileName.substring(0, inputFileName.lastIndexOf('.')) + ".xlsx";
        Path outputFile = Paths.get(outputPath, outputFileName);

        List<String[]> records = new ArrayList<>();

        String[] headers = null;

        try (BufferedReader reader = Files.newBufferedReader(inputFile)) {
            String line;
            int lineNumber = 0;

            while ((line = reader.readLine()) != null) {
                line = line.trim();
                if (line.isEmpty()) continue;

                if (lineNumber == 0) {
                    // Process header
                    String[] rawHeaders = line.split("\\|", -1);  // preserve empty strings
                    List<String> cleanedHeaderList = new ArrayList<>();
                    for (String h : rawHeaders) {
                        h = h.trim();
                        // Skip known output-calculated columns
                        String hLower = h.toLowerCase();
                        if (h.isEmpty() ||
                                hLower.contains("available cradit balence") ||
                                hLower.contains("cradit utlization") ||
                                hLower.contains("available credit balance") ||
                                hLower.contains("credit utilization")) {
                            continue;
                        }
                        cleanedHeaderList.add(h);
                    }
                    headers = cleanedHeaderList.toArray(new String[0]);
                } else {
                    // Process data
                    String[] rawFields = line.split("\\|", -1);
                    List<String> cleanedFields = new ArrayList<>();
                    for (String f : rawFields) {
                        f = f.trim();
                        if (!f.isEmpty()) {
                            cleanedFields.add(f);
                        }
                    }
                    records.add(cleanedFields.toArray(new String[0]));
                }
                lineNumber++;
            }

            if (headers == null) {
                throw new IllegalArgumentException("DAT file does not contain header line.");
            }

            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Accounts");

                Row headerRow = sheet.createRow(0);
                int colCount = headers.length + 2;

                for (int i = 0; i < headers.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers[i]);
                }

                headerRow.createCell(headers.length).setCellValue("Available credit balance");
                headerRow.createCell(headers.length + 1).setCellValue("Credit utilization");

                int creditLimitIdx = -1;
                int spendingIdx = -1;
                for (int i = 0; i < headers.length; i++) {
                    String h = headers[i].toLowerCase();
                    if (h.contains("cradit limit") || h.contains("credit limit")) creditLimitIdx = i;
                    if (h.contains("spending")) spendingIdx = i;
                }

                if (creditLimitIdx == -1 || spendingIdx == -1) {
                    throw new IllegalArgumentException("Required columns 'Credit Limit' or 'Spending' not found in header.");
                }

                CellStyle twoDecimalStyle = workbook.createCellStyle();
                DataFormat format = workbook.createDataFormat();
                twoDecimalStyle.setDataFormat(format.getFormat("0.00"));

                int rowNum = 1;
                for (String[] record : records) {
                    Row row = sheet.createRow(rowNum++);

                    for (int i = 0; i < headers.length; i++) {
                        Cell cell = row.createCell(i);
                        if (i < record.length) {
                            cell.setCellValue(record[i]);
                        } else {
                            cell.setCellValue("");
                        }
                    }

                    double creditLimit = 0;
                    double spending = 0;
                    try {
                        creditLimit = Double.parseDouble(record[creditLimitIdx].replaceAll(",", ""));
                    } catch (Exception e) {
                        creditLimit = 0;
                    }
                    try {
                        spending = Double.parseDouble(record[spendingIdx].replaceAll(",", ""));
                    } catch (Exception e) {
                        spending = 0;
                    }

                    double availableBalance = creditLimit - spending;
                    double utilization = creditLimit == 0 ? 0 : (spending / creditLimit) * 100;

                    Cell availableBalanceCell = row.createCell(headers.length);
                    availableBalanceCell.setCellValue(availableBalance);
                    availableBalanceCell.setCellStyle(twoDecimalStyle);

                    Cell utilizationCell = row.createCell(headers.length + 1);
                    utilizationCell.setCellValue(utilization);
                    utilizationCell.setCellStyle(twoDecimalStyle);
                }

                for (int i = 0; i < colCount; i++) {
                    sheet.autoSizeColumn(i);
                }

                try (FileOutputStream fos = new FileOutputStream(outputFile.toFile())) {
                    workbook.write(fos);
                }
            }
        }
    }
}
