package io.file.parser.xlsx;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import static io.gupshup.goldengate.async.fileprocessing.constants.Constants.TEMPORARY_FILE_NAME;
import static io.gupshup.goldengate.async.fileprocessing.constants.Constants.XLSX_EXTENSION;

public class XSSFParser {
    private static final Logger logger = LoggerFactory.getLogger(XSSFParser.class);

    /**
     * Uses the XSSF Event SAX helpers to do most of the work
     * of parsing the Sheet XML, and outputs the contents
     * as a (basic) CSV.
     */
    private class SheetToCSV implements SheetContentsHandler {
        /**
         * Number of columns to read starting with leftmost
         */
        private int minColumns = -1;
        private int numberOfHeaders = -1;
        private boolean firstCellOfRow;
        private int currentRow = -1;
        private int currentCol = -1;

        private void outputMissingRows(int number) {
            for (int i = 0; i < number; i++) {
                for (int j = 0; j < minColumns; j++) {
                    output.append(',');
                }
                output.append('\n');
            }
        }

        @Override
        public void startRow(int rowNum) {
            // If there were gaps, output the missing rows
            outputMissingRows(rowNum - currentRow - 1);
            // Prepare for this row
            firstCellOfRow = true;
            currentRow = rowNum;
            currentCol = -1;
        }

        @Override
        public void endRow(int rowNum) {
            if (rowNum == 0) {
                numberOfHeaders = currentCol + 1;
            }
            minColumns = Math.max(minColumns, numberOfHeaders - 1);
            // Ensure the minimum number of columns
            for (int i = currentCol; i < minColumns; i++) {
                output.append(',');
            }
            output.append('\n');
        }

        @Override
        public void cell(String cellReference, String formattedValue,
                         XSSFComment comment) {
            if (firstCellOfRow) {
                firstCellOfRow = false;
            } else {
                output.append(',');
            }

            // gracefully handle missing CellRef here in a similar way as XSSFCell does
            if (cellReference == null) {
                cellReference = new CellAddress(currentRow, currentCol).formatAsString();
            }

            // Did we miss any cells?
            int thisCol = (new CellReference(cellReference)).getCol();
            int missedCols = thisCol - currentCol - 1;
            for (int i = 0; i < missedCols; i++) {
                output.append(',');
            }

            // no need to append anything if we do not have a value
            if (formattedValue == null) {
                return;
            }

            currentCol = thisCol;

            // Number or string?
            try {
                //noinspection ResultOfMethodCallIgnored
                Double.parseDouble(formattedValue);
                output.append(formattedValue);
            } catch (Exception e) {
                // remove quotes if they are already there
                if (formattedValue.startsWith("\"") && formattedValue.endsWith("\"")) {
                    formattedValue = formattedValue.substring(1, formattedValue.length()-1);
                }
                output.append('"');
                // encode double-quote with two double-quotes to produce a valid CSV format
                String updatedValue = formattedValue.replace("\"", "\"\"");
                output.append(updatedValue);
                output.append('"');
            }
        }
    }

    private final OPCPackage xlsxPackage;

    /**
     * Destination for data
     */
    private final PrintStream output;

    /**
     * Creates a new XLSX ->; CSV converter
     *
     * @param pkg    The XLSX package to process
     * @param output The PrintStream to output the CSV to
     */
    public XSSFParser(OPCPackage pkg, PrintStream output) {
        this.xlsxPackage = pkg;
        this.output = output;
    }

    /**
     * Parses and shows the content of one sheet
     * using the specified styles and shared-strings tables.
     *
     * @param styles           The table of styles that may be referenced by cells in the sheet
     * @param strings          The table of strings that may be referenced by cells in the sheet
     * @param sheetInputStream The stream to read the sheet-data from.
     * @throws IOException  An IO exception from the parser,
     *                      possibly from a byte stream or character stream
     *                      supplied by the application.
     * @throws SAXException if parsing the XML data fails.
     */
    private static void processSheet(
            Styles styles,
            SharedStrings strings,
            SheetContentsHandler sheetHandler,
            InputStream sheetInputStream) throws IOException, SAXException {
        // set emulateCSV=true on DataFormatter - it is also possible to provide a Locale
        // when POI 5.2.0 is released, you can call formatter.setUse4DigitYearsInAllDateFormats(true)
        // to ensure all dates are formatted with 4 digit years
        DataFormatter formatter = new DataFormatter(true);
        formatter.setUse4DigitYearsInAllDateFormats(true);
        formatter.addFormat("General", new java.text.DecimalFormat("#.##############################"));//dataformater made to work like excel for large numbers
        formatter.addFormat("@",new java.text.DecimalFormat("#.##############################"));//handling large number in text format
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = XMLHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch (ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        }
    }


    /**
     * Initiates the processing of the XLS workbook file to CSV.
     *
     * @throws IOException  If reading the data from the package fails.
     * @throws SAXException if parsing the XML data fails.
     */
    private void process() throws IOException, OpenXML4JException, SAXException {
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        try (InputStream stream = iter.next()) {
            String sheetName = iter.getSheetName();

            try {
                processSheet(styles, strings, new SheetToCSV(), stream);
            } catch (NumberFormatException e) {
                throw new IOException("Failed to parse sheet " + sheetName, e);
            }
        }
    }

    private static ByteArrayInputStream xlsxToCsv(InputStream inputStream) throws IOException, OpenXML4JException, SAXException {
        long startTime = System.currentTimeMillis();
        logger.info("###In xlsxToCsv(), xlsx to csv conversion starts at: {}", startTime);

        ByteArrayInputStream byteArrayInputStream;
        Path tempFile = Files.createTempFile(TEMPORARY_FILE_NAME, XLSX_EXTENSION);
        Files.copy(inputStream, tempFile, StandardCopyOption.REPLACE_EXISTING);

        // The package open is instantaneous, as it should be.
        try (OPCPackage opcPackage = OPCPackage.open(tempFile.toString(), PackageAccess.READ);
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
             PrintStream out = new PrintStream(outputStream)) {

            XSSFParser xssfParser = new XSSFParser(opcPackage, out);
            xssfParser.process();
            out.flush();
            byteArrayInputStream = new ByteArrayInputStream(outputStream.toByteArray());
        }
        logger.info("###Time taken for .xlsx to .csv conversion: {} ms", System.currentTimeMillis() - startTime);
        return byteArrayInputStream;
    }

    private static void streamToCsv(ByteArrayInputStream byteArrayInputStream) {
        try {
            InputStreamReader isr = new InputStreamReader(byteArrayInputStream);
            BufferedReader br = new BufferedReader(isr);
            
            PrintWriter writer = new PrintWriter(new FileWriter("output.csv"));

            String line;
            while ((line = br.readLine()) != null) {
                writer.println(line);
            }
            
            br.close();
            writer.close();
            logger.info("CSV file created successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static void main(String[] args) throws IOException, OpenXML4JException, SAXException {
        FileInputStream inputStream = new FileInputStream("test.xlsx");
        ByteArrayInputStream byteArrayInputStream = xlsxToCsv(inputStream);
        streamToCsv(byteArrayInputStream);
    }

}
