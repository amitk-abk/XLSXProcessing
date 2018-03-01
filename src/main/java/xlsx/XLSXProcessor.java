package xlsx;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class XLSXProcessor {

    public static void main(String[] args) {
        XLSXProcessor xlsxProcessor = new XLSXProcessor();
        try {
            List<Parameters> parameters = xlsxProcessor.parseXlsxFile(new File("\\resources\\InputFile.xlsx"));	//Path to input file
            System.out.println(parameters);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private List<Parameters> parseXlsxFile(File inputFile) throws IOException {
        OPCPackage container = null;
        List<Parameters> customParams = null;
        try {
            container = OPCPackage.open(inputFile.getAbsolutePath());
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(container);
            XSSFReader xssfReader = new XSSFReader(container);
            StylesTable styles = xssfReader.getStylesTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();

            while (iter.hasNext()) {
                InputStream stream = iter.next();
                customParams = processSheet(styles, strings, stream);
                stream.close();
                stream = null;
                break; // Just process first sheet and return
            }
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (SAXException e) {
            e.printStackTrace();
        } catch (OpenXML4JException e) {
            e.printStackTrace();
        } finally {
            if (container != null) {
                try {
                    container.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
                container = null;
            }
        }
        return customParams;
    }

    private List<Parameters> processSheet(StylesTable styles, ReadOnlySharedStringsTable strings, InputStream sheetInputStream) throws SAXException, IOException {

        InputSource sheetSource = new InputSource(sheetInputStream);
        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        try {
            SAXParser saxParser = saxFactory.newSAXParser();
            XMLReader sheetParser = saxParser.getXMLReader();
            List<Parameters> customParams = new ArrayList<>();
            List<RowModel> dataList = new ArrayList<>();
            ContentHandler handler = new XSSFSheetXMLHandler(styles, strings,
                    new SheetProcessorClass(dataList, customParams),
                    false // means result instead of formula
            );
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
            dataList.clear();
            return customParams;
        } catch (ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage(), e);
        }
    }

    class Parameters {
        Long mobile;
        String name;
        String city;
        String country;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getCity() {
            return city;
        }

        public void setCity(String city) {
            this.city = city;
        }

        public String getCountry() {
            return country;
        }

        public void setCountry(String country) {
            this.country = country;
        }

        public Long getMobile() {
            return mobile;
        }

        public void setMobile(Long mobile) {
            this.mobile = mobile;
        }

        @Override
        public String toString() {
            final StringBuilder sb = new StringBuilder("Parameters{");
            sb.append("mobile=").append(mobile);
            sb.append(", name='").append(name).append('\'');
            sb.append(", city='").append(city).append('\'');
            sb.append(", country='").append(country).append('\'');
            sb.append('}');
            return sb.toString();
        }
    }

    class RowModel {
        private int rowNum;
        private List<String> rowDataList = new ArrayList<>();
        private Map<String, String> rowMap = new HashMap<>();

        public int getRowNum() {
            return rowNum;
        }

        public void setRowNum(int rowNum) {
            this.rowNum = rowNum;
        }

        public List<String> getRowDataList() {
            return rowDataList;
        }

        public void setRowDataList(List<String> rowDataList) {
            this.rowDataList = rowDataList;
        }

        public Map<String, String> getRowMap() {
            return rowMap;
        }

        public void setRowMap(Map<String, String> rowMap) {
            this.rowMap = rowMap;
        }
    }

    private final class SheetProcessorClass implements SheetContentsHandler {

        private List<RowModel> dataList = null;
        private List<Parameters> customParams = null;
        private Parameters parameter;
        private RowModel rowData;


        public SheetProcessorClass(List<RowModel> dataList, List<Parameters> customParams) {
            this.customParams = customParams;
            this.dataList = dataList;
        }

        @Override
        public void startRow(int rowNum) {
            parameter = new Parameters();
            rowData = new RowModel();
            rowData.setRowNum(rowNum);
        }

        @Override
        public void endRow(int args) {
            // Process first row for headers in xlsx file
            if (args == 0)
                dataList.add(rowData);

            // Process data rows from xlsx file.
            if (args >= 1)
                processRow();
        }

        @Override
        public void cell(String cellHeader, String cellData, XSSFComment comment) {
            cellHeader = parse(cellHeader);
            rowData.getRowMap().put(cellHeader, cellData);
        }

        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {
        }

        private void processRow() {
            RowModel headers = dataList.get(0);
            Map<String, String> headerMap = headers.getRowMap();
            for (String key : headerMap.keySet()) {
                String value = rowData.getRowMap().get(key);
                if (key.equalsIgnoreCase("A") && rowData.getRowNum() > 0) {
                    parameter.setMobile(Long.valueOf(value));
                }
                if (key.equalsIgnoreCase("B") && rowData.getRowNum() > 0) {
                    parameter.setName(value);
                }
                if (key.equalsIgnoreCase("C") && rowData.getRowNum() > 0) {
                    parameter.setCity(value);
                }
                if (key.equalsIgnoreCase("D") && rowData.getRowNum() > 0) {
                    parameter.setCountry(value);
                }
            }

            if (rowData.getRowNum() > 0) {
                customParams.add(parameter);
            }
        }

        private String parse(String str) {
            Matcher match = Pattern.compile("[0-9]+|[a-z]+|[A-Z]+").matcher(str);
            while (match.find()) {
                String subGroup = match.group();
                if (subGroup.matches("[a-zA-Z]+"))
                    return subGroup;
            }
            return str;
        }
    }
}
