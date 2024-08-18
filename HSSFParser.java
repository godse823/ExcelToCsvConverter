package io.file.parser.xls;

import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder.SheetRecordCollectingListener;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class HSSFParser implements HSSFListener {
    private static final Logger logger = LoggerFactory.getLogger(HSSFParser.class);

    private final POIFSFileSystem fs;
    private final PrintStream output;
    private int lastRowNumber;
    private int lastColumnNumber;

    /**
     * Should we output the formula, or the value it has?
     */
    private final boolean outputFormulaValues = true;

    /**
     * For parsing Formulas
     */
    private SheetRecordCollectingListener workbookBuildingListener;
    private HSSFWorkbook stubWorkbook;

    // Records we pick up as we process
    private SSTRecord sstRecord;
    private FormatTrackingHSSFListener formatListener;

    /**
     * So we known which sheet we're on
     */
    private int sheetIndex = -1;
    private BoundSheetRecord[] orderedBSRs;
    private final List<BoundSheetRecord> boundSheetRecords = new ArrayList<>();

    // For handling formulas with string results
    private int nextRow;
    private int nextColumn;
    private boolean outputNextStringRecord;
    private int minColumns = -1;
    private int numberOfHeaders = 0;

    /**
     * Creates a new XLS -&gt; CSV converter
     *
     * @param fs     The POIFSFileSystem to process
     * @param output The PrintStream to output the CSV to
     */
    public HSSFParser(POIFSFileSystem fs, PrintStream output) {
        this.fs = fs;
        this.output = output;
    }

    /**
     * Creates a new XLS -&gt; CSV converter
     *
     * @param inputStream The stream to process
     * @throws IOException if the file cannot be read or parsing the file fails
     */
    public HSSFParser(InputStream inputStream, PrintStream output) throws IOException {
        this(
                new POIFSFileSystem(inputStream),
                output
        );
    }

    /**
     * Initiates the processing of the XLS file to CSV
     *
     * @throws IOException if the workbook contained errors
     */
    private void process() throws IOException {
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        formatListener = new FormatTrackingHSSFListener(listener);//is responsible for tracking the format of cells.

        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();

        if (outputFormulaValues) {
            request.addListenerForAllRecords(formatListener);
        } else {
            workbookBuildingListener = new SheetRecordCollectingListener(formatListener);
            request.addListenerForAllRecords(workbookBuildingListener);
        }
        //event-driven parsing of the Excel file
        factory.processWorkbookEvents(request, fs);
    }

    /**
     * Main HSSFListener method, processes events, and outputs the
     * CSV as the file is processed.
     */
    @Override
    public void processRecord(Record record) {
        int thisRow = -1;
        int thisColumn = -1;
        String thisStr = null;
        //sid : Stream ID, and it is a unique identifier associated with each record in a Microsoft Excel file
        /*  Record can be :
         * Cell Records : BlankRecord, BoolErrRecord, FormulaRecord, LabelRecord, NumberRecord, RKRecord, etc
         * Row and Column Records: RowRecord, ColumnInfoRecord
         * Sheet Records: BoundSheetRecord
         * File-Level Records: BOFRecord
         */
        if(sheetIndex<1) {
            switch (record.getSid()) {
                case BoundSheetRecord.sid:
                    boundSheetRecords.add((BoundSheetRecord) record);
                    break;
                case BOFRecord.sid://Beginning of File Record
                    BOFRecord br = (BOFRecord) record;
                    if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
                        // Create sub workbook if required
                        if (workbookBuildingListener != null && stubWorkbook == null) {
                            stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
                        }

                        // Output the worksheet name
                        // Works by ordering the BSRs by the location of their BOFRecords, and then knowing that we process BOFRecords in byte offset order
                        sheetIndex++;
                        if (orderedBSRs == null) {
                            orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
                        }
                    }
                    break;

                case SSTRecord.sid:
                    //shared string table in the workbook.
                    sstRecord = (SSTRecord) record;
                    break;

                case BlankRecord.sid:
                    //Handling BlankRecord
                    BlankRecord brec = (BlankRecord) record;

                    thisRow = brec.getRow();
                    thisColumn = brec.getColumn();
                    thisStr = "";
                    break;
                case BoolErrRecord.sid:
                    //Handling BoolErrRecord
                    BoolErrRecord berec = (BoolErrRecord) record;

                    thisRow = berec.getRow();
                    thisColumn = berec.getColumn();
                    thisStr = "";
                    break;

                case FormulaRecord.sid:
                    //it extracts the row and column information.
                    FormulaRecord frec = (FormulaRecord) record;

                    thisRow = frec.getRow();
                    thisColumn = frec.getColumn();

                    if (outputFormulaValues) {
                        if (Double.isNaN(frec.getValue())) {
                            // Formula result is a string. This is stored in the next record
                            outputNextStringRecord = true;
                            nextRow = frec.getRow();
                            nextColumn = frec.getColumn();
                        } else {
                            thisStr = formatListener.formatNumberDateCell(frec);
                        }
                    } else {
                        thisStr = HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression());
                    }
                    break;
                case StringRecord.sid:
                    if (outputNextStringRecord) {
                        // String for formula
                        StringRecord srec = (StringRecord) record;
                        thisStr = srec.getString();
                        thisRow = nextRow;
                        thisColumn = nextColumn;
                        outputNextStringRecord = false;
                    }
                    break;

                case LabelRecord.sid:
                    LabelRecord lrec = (LabelRecord) record;

                    thisRow = lrec.getRow();
                    thisColumn = lrec.getColumn();
                    thisStr = lrec.getValue();
                    break;
                case LabelSSTRecord.sid:
                    LabelSSTRecord lsrec = (LabelSSTRecord) record;

                    thisRow = lsrec.getRow();
                    thisColumn = lsrec.getColumn();
                    if (sstRecord == null) {
                        thisStr = "";
                    } else {
                        thisStr = sstRecord.getString(lsrec.getSSTIndex()).toString();
                    }
                    break;
                case NumberRecord.sid:
                    NumberRecord numrec = (NumberRecord) record;

                    thisRow = numrec.getRow();
                    thisColumn = numrec.getColumn();

                    String formattedValue;
                    double rawValue = numrec.getValue();

                    if (rawValue == Math.floor(rawValue)) {
                        // If the value is an integer, format it as a string without decimal places
                        formattedValue = String.format("%.0f", rawValue);
                    } else {
                        // If the value has decimal places, format it normally
                        formattedValue = String.valueOf(rawValue);
                    }

                    thisStr = formattedValue;
                    break;
                default:
                    break;
            }

            if (thisRow == 0) {
                numberOfHeaders++;
                minColumns = Math.max(minColumns, numberOfHeaders - 1);
            }
            // Handle new row
            if (thisRow != -1 && thisRow != lastRowNumber) {
                lastColumnNumber = -1;
            }

            // Handle missing column
            if (record instanceof MissingCellDummyRecord) {
                MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
                thisRow = mc.getRow();
                thisColumn = mc.getColumn();
                thisStr = "";
            }

            // If we got something to print out, do so
            if (thisStr != null) {
                if (thisColumn > 0) {
                    output.print(',');
                }
                // remove quotes if they are already there
                if (thisStr.startsWith("\"") && thisStr.endsWith("\"")) {
                    thisStr = thisStr.substring(1, thisStr.length()-1);
                }
                output.append('"');
                // encode double-quote with two double-quotes to produce a valid CSV format
                String updatedValue = thisStr.replace("\"", "\"\"");
                output.print(updatedValue);
                output.append('"');
            }

            // Update column and row count
            if (thisRow > -1)
                lastRowNumber = thisRow;
            if (thisColumn > -1)
                lastColumnNumber = thisColumn;

            // Handle end of row
            if (record instanceof LastCellOfRowDummyRecord) {
                // Print out any missing commas if needed
                if (minColumns > 0) {
                    // Columns are 0 based
                    if (lastColumnNumber == -1) {
                        lastColumnNumber = 0;
                    }
                    for (int i = lastColumnNumber; i < (minColumns); i++) {
                        output.print(',');
                    }
                }

                // We're onto a new row
                lastColumnNumber = -1;

                // End the row
                output.println();
            }
        }
    }

    private static ByteArrayInputStream xlsToCsv(InputStream inputStream) throws IOException {
        long startTime = System.currentTimeMillis();
        logger.info("###In xlsToCsv(), xls to csv conversion starts at: {}", startTime);

        ByteArrayInputStream byteArrayInputStream;

        // The package open is instantaneous, as it should be.
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
             PrintStream out = new PrintStream(outputStream)) {
            HSSFParser hssfParser = new HSSFParser(inputStream, out);
            hssfParser.process();
            out.flush();
            byteArrayInputStream = new ByteArrayInputStream(outputStream.toByteArray());
        }
        logger.info("###Time taken for .xls to .csv conversion: {} ms", System.currentTimeMillis() - startTime);
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
        FileInputStream inputStream = new FileInputStream("test.xls");
        ByteArrayInputStream byteArrayInputStream = xlsToCsv(inputStream);
        streamToCsv(byteArrayInputStream);
    }

}
