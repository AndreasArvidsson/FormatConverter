package andreas.arvidsson.csvtospreadsheet;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 *
 * @author Andreas Arvidsson
 */
public class Converter {

    private final String sheetStartDelimiter = "[";
    private final String sheetEndDelimiter = "]";

    public void run(File source, File destination, String delimiter) throws FileNotFoundException, IOException {
        //Use streaming workbook to support really large files.
        Workbook wb = new SXSSFWorkbook(100);
        try (BufferedReader br = new BufferedReader(new FileReader(source))) {
            Sheet sheet = null;
            Row row;
            Cell cell;
            int sheetIndex = 1;
            int rowIndex = 0;
            String line;
            while ((line = br.readLine()) != null) {
                if (line.startsWith(sheetStartDelimiter) && line.endsWith(sheetEndDelimiter)) {
                    String sheetName = line.substring(sheetStartDelimiter.length(), line.length() - sheetEndDelimiter.length()).trim();
                    if (sheetName.isEmpty()) {
                        sheetName = "Sheet" + sheetIndex++;
                    }
                    sheet = wb.createSheet(sheetName);
                    rowIndex = 0;
                }
                else {
                    if (sheet == null) {
                        sheet = wb.createSheet("Sheet" + sheetIndex++);
                        rowIndex = 0;
                    }
                    row = sheet.createRow(rowIndex++);
                    String[] columns = line.split(delimiter);
                    for (int i = 0; i < columns.length; ++i) {
                        cell = row.createCell(i);
                        cell.setCellValue(columns[i]);
                    }
                }
            }
        }
        save(wb, destination);
    }

    private void save(Workbook wb, File destination) throws FileNotFoundException, IOException {
        try (FileOutputStream fileOut = new FileOutputStream(destination)) {
            wb.write(fileOut);
        }
    }

}
