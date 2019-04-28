package andreas.arvidsson.spreadsheettojson;

import com.fasterxml.jackson.core.JsonEncoding;
import com.fasterxml.jackson.core.JsonFactory;
import com.fasterxml.jackson.core.JsonGenerator;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Andreas Arvidsson
 */
public class Converter {

    public void run(File source, File destination, boolean addSheetName) throws FileNotFoundException, IOException {
        Workbook wb;
        try (InputStream is = new FileInputStream(source)) {
            //XSSF (.xlsx)
            if (source.getName().toLowerCase().endsWith(".xlsx")) {
                wb = new XSSFWorkbook(is);
            }
            //HSSF (.xls)
            else if (source.getName().toLowerCase().endsWith(".xls")) {
                wb = new HSSFWorkbook(is);
            }
            else {
                throw new RuntimeException("Unnsuported file ending " + source.getName());
            }
            try (JsonGenerator jGenerator = new JsonFactory().createGenerator(destination, JsonEncoding.UTF8)) {
                jGenerator.writeStartArray();
                jGenerator.writeRaw("\n");
                for (Sheet sheet : wb) {
                    parseSheet(jGenerator, sheet, addSheetName);
                }
                jGenerator.writeEndArray();
            }
        }
    }

    private void parseSheet(JsonGenerator jGenerator, Sheet sheet, boolean addSheetName) throws IOException {
        final List<String> columnNames = getColumnNames(sheet);
        for (int i = 1; i <= sheet.getLastRowNum(); ++i) {
            final Row row = sheet.getRow(i);
            jGenerator.writeStartObject();
            if (addSheetName) {
                jGenerator.writeStringField("SheetName", sheet.getSheetName());
            }
            for (int j = 0; j < columnNames.size(); ++j) {
                final Cell cell = row.getCell(j);
                final String value = cell != null ? cell.toString() : null;
                jGenerator.writeStringField(columnNames.get(j), value);
            }
            jGenerator.writeEndObject();
            jGenerator.writeRaw("\n");
        }
    }

    private List<String> getColumnNames(Sheet sheet) {
        final List<String> res = new ArrayList();
        final Row row = sheet.getRow(0);
        if (row != null) {
            for (Cell cell : row) {
                res.add(cell.toString());
            }
        }
        return res;
    }

}
