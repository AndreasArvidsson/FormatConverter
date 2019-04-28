package andreas.arvidsson.csvtojson;

import com.fasterxml.jackson.core.JsonEncoding;
import com.fasterxml.jackson.core.JsonFactory;
import com.fasterxml.jackson.core.JsonGenerator;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;

/**
 *
 * @author Andreas Arvidsson
 */
public class Converter {

    public void run(File source, File destination, String delimiter) throws FileNotFoundException, IOException {
        try (BufferedReader br = new BufferedReader(new FileReader(source))) {
            try (JsonGenerator jGenerator = new JsonFactory().createGenerator(destination, JsonEncoding.UTF8)) {
                final String[] columnNames = br.readLine().split(delimiter);

                jGenerator.writeStartArray();
                jGenerator.writeRaw("\n");

                String line;
                while ((line = br.readLine()) != null) {
                    parseLine(jGenerator, columnNames, line.split(delimiter));
                }

                jGenerator.writeEndArray();
            }
        }
    }

    private void parseLine(JsonGenerator jGenerator, String[] columnNames, String[] values) throws IOException {
        jGenerator.writeStartObject();
        for (int i = 0; i < columnNames.length; ++i) {
            jGenerator.writeStringField(columnNames[i], i < values.length ? values[i] : null);
        }
        jGenerator.writeEndObject();
        jGenerator.writeRaw("\n");
    }

}
