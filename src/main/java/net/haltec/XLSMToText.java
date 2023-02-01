package net.haltec;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Map;
import java.util.regex.Pattern;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.macros.VBAMacroReader;
import org.apache.poi.ss.usermodel.DataFormatter;
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
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

public class XLSMToText {
    static final int MIN_CELL_WIDTH = 40;

    private class OutCell {
        public int col;
        public String[] text_rows;

        public int width = -1;

        public OutCell(int col, String text) {
            this.col = col;
            this.text_rows = text.split("\n");
            for (String t : this.text_rows) {
                width = Math.max(width, t.length());
            }
            width = Math.max(width, MIN_CELL_WIDTH);
        }
    }
    private class SheetToText implements SheetContentsHandler {
        private ArrayList<OutCell> row_cells;
        int last_row = -1;

        @Override
        public void startRow(int rowNum) {
            row_cells = new ArrayList<>();

        }

        @Override
        public void endRow(int rowNum) {
            if (row_cells.isEmpty()) {
                return;
            }
            boolean skipped_rows = rowNum != last_row + 1;
            last_row = rowNum;

            int last_col = -1;
            for (OutCell cell : row_cells) {
                if (cell.col > 0) {
                    if (cell.col == last_col + 1) {
                        output.print(skipped_rows ? "╤" : "┬");
                    } else {
                        output.print(skipped_rows ? "╦" : "╥");
                    }
                }
                output.print(skipped_rows ? "═" : "─");
                output.print(" ");
                String loc = rowColToExcel(rowNum, cell.col);
                output.print(loc);
                output.print(" ");

                // plus one space at the start, plus one space at the end
                // minus two spaces at the loc text start and end, minus the first line char before the loc text
                output.print((skipped_rows ? "═" : "─").repeat(cell.width + 2 - 3 - loc.length()));
                last_col = cell.col;
            }
            output.print("\n");

            boolean all_cells_done;
            int text_row = 0;
            do {
                all_cells_done = true;
                last_col = -1;
                int counter = 0;
                for (OutCell cell : row_cells) {
                    if (cell.col > 0) {
                        if (cell.col == last_col + 1) {
                            output.print("│");
                        } else {
                            output.print("║");
                        }
                    }

                    if (text_row < cell.text_rows.length && !cell.text_rows[text_row].isEmpty()) {
                        all_cells_done = false;
                        output.print(" ");
                        output.print(cell.text_rows[text_row]);
                        if (counter < row_cells.size() - 1) {
                            output.print(" ".repeat(cell.width - cell.text_rows[text_row].length() + 1));
                        }
                    }
                    else {
                        if (counter < row_cells.size() - 1) {
                            output.print(" ".repeat(cell.width + 2));
                        }
                    }
                    last_col = cell.col;
                    counter++;
                }
                output.print("\n");
                text_row++;
            } while (!all_cells_done);
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            if(cellReference != null) {
                row_cells.add(new OutCell((new CellReference(cellReference)).getCol(), formattedValue));
            }
        }
    }

    private final OPCPackage xlsxPackage;

    private final PrintStream output;

    /**
     * Creates a new XLSX -&gt; CSV converter
     *
     * @param pkg        The XLSX package to process
     * @param output     The PrintStream to output the CSV to
     */
    public XLSMToText(OPCPackage pkg, PrintStream output) {
        this.xlsxPackage = pkg;
        this.output = output;
    }

    /**
     * Parses and shows the content of one sheet
     * using the specified styles and shared-strings tables.
     *
     * @param styles The table of styles that may be referenced by cells in the sheet
     * @param strings The table of strings that may be referenced by cells in the sheet
     * @param sheetInputStream The stream to read the sheet-data from.

     * @throws java.io.IOException An IO exception from the parser,
     *            possibly from a byte stream or character stream
     *            supplied by the application.
     * @throws SAXException if parsing the XML data fails.
     */
    public void processSheet(Styles styles, SharedStrings strings, SheetContentsHandler sheetHandler, InputStream sheetInputStream) throws IOException, SAXException {
        DataFormatter formatter = new DataFormatter(true);
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = XMLHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, null, strings, sheetHandler, formatter, true);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch(ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        }
    }

    /**
     * Initiates the processing of the XLS workbook file to CSV.
     *
     * @throws IOException If reading the data from the package fails.
     * @throws SAXException if parsing the XML data fails.
     */
    public void extractSheets() throws IOException, OpenXML4JException, SAXException {
        output.println("Sheets {");
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        int index = 0;
        while (iter.hasNext()) {
            try (InputStream stream = iter.next()) {
                String sheetName = iter.getSheetName();
                this.output.println();
                this.output.println("sheet " + sheetName + " {");
                processSheet(styles, strings, new SheetToText(), stream);
                this.output.println("}");
            }
            ++index;
        }
        output.println("}");
    }

    public static void main(String[] args) throws Exception {
        if (args.length < 1) {
            System.err.println("Use:");
            System.err.println("  XLSX2CSV <xlsx file> [min columns]");
            return;
        }

        PrintStream csvOutStream;
        if (args.length > 1) {
            csvOutStream = new PrintStream(args[0] + ".csv", StandardCharsets.UTF_8);
        }
        else {
            csvOutStream = System.out;
        }

        File xlsxFile = new File(args[0]);
        if (!xlsxFile.exists()) {
            System.err.println("Not found or not a file: " + xlsxFile.getPath());
            return;
        }

        OPCPackage p = OPCPackage.open(xlsxFile.getPath(), PackageAccess.READ);
        try {
            XLSMToText xlsx2csv = new XLSMToText(p, csvOutStream);
            xlsx2csv.extractSheets();
            xlsx2csv.extractMacros(xlsxFile);
        } finally {
            p.revert();
        }


        csvOutStream.close();
    }

    public static String rowColToExcel(int row, int col) {
        // We have to work with 0-based numbers (0=A, 25=Z), otherwise the modulo operator won't work.
        // (26 / 26 = 1 Remainder 0 and not 0 Remainder 26 which would lead to A0 and not Z)

        String alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        if (row < 0 || col < 0) {
            throw new IndexOutOfBoundsException("Row and column must both be greater than 0, got row = " + row + ", col = " + col);
        }

        String result = "";

        int remainder;

        while ( col >= 0 ) {
            remainder = col % alphabet.length();
            result = alphabet.substring(remainder, remainder + 1) + result;

            // - 1, to get back to a 0-based system.
            // 1 Remainder 0 should result in A(=0).
            // 0 Remainder 0 should trigger loop end.

            col = col / alphabet.length() - 1;
        }

        return result + (row+1);
    }

    private void extractMacros(File input) {
        output.println("Macros {");
        Pattern vb_base_pattern = Pattern.compile("^Attribute VB_Base = \"0\\{[^}]+\\}\\{[^}]+\\}\"$", Pattern.MULTILINE);
        Pattern vb_function_pattern = Pattern.compile("((?:Private|Public|) (?:Function|Sub) )([^(]+\\([^)]*\\).*)$", Pattern.MULTILINE);
        try (VBAMacroReader reader = new VBAMacroReader(input)) {
            final Map<String, String> macros = reader.readMacros();
            for (Map.Entry<String, String> entry : macros.entrySet()) {
                String moduleName = entry.getKey();
                String moduleCode = entry.getValue();
                moduleCode = moduleCode.replace("\r\n", "\n");
                moduleCode = vb_base_pattern.matcher(moduleCode).replaceAll("Attribute VB_Base = \"0{XXX}{XXX}\"");
                moduleCode = vb_function_pattern.matcher(moduleCode).replaceAll("$1" + moduleName + "::$2");
                output.println("module " + moduleName + " {");
                output.println(moduleCode);
                output.println("}");
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        output.println("}");
    }
}
