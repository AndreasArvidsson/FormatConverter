package andreas.arvidsson.spreadsheettojson;

import java.io.File;
import java.io.IOException;

/**
 *
 * @author Andreas Arvidsson
 */
public class Main {

    public static String fileEnding = ".json";

    public static void main(String args[]) throws IOException {
        if (args.length < 2) {
            throw new RuntimeException("File source{0} and destination{1} is required!");
        }
        final Converter converter = new Converter();
        final File source = new File(args[0]);
        final File destination = getDestination(args[1]);
        final boolean addSheetName = addSheetName(args);
        converter.run(source, destination, addSheetName);
    }

    private static File getDestination(String value) {
        //Already has correct file ending.
        if (value.endsWith(fileEnding)) {
            return new File(value);
        }
        //Need to add ending.
        return new File(value + fileEnding);
    }

    private static boolean addSheetName(String args[]) {
        if (args.length > 2) {
            return args[2].toLowerCase().equals("addsheetname");
        }
        return false;
    }

}
