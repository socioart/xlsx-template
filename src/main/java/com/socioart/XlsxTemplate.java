package com.socioart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import java.io.*;
import java.util.Iterator;
import java.util.regex.Pattern;
import java.util.regex.Matcher;
import org.json.*;

/**
 * Hello world!
 *
 */
public class XlsxTemplate
{
    public static void main( String[] args ) throws IOException
    {
        FileInputStream input_fis;
        FileInputStream json_fis;

        String input = args[0];
        String output = args[1];
        String json_file = args[2];
        Pattern pattern = Pattern.compile("\\{\\{(.*)\\}\\}");

        try {
            input_fis = new FileInputStream(input);
        } catch(FileNotFoundException e) {
            System.err.println( "Templatel file not found" );
            System.exit(1);
            return;
        }

        try {
            json_fis = new FileInputStream(json_file);
        } catch(FileNotFoundException e) {
            System.err.println( "Data file not found" );
            System.exit(1);
            input_fis.close();
            return;
        }

        JSONObject variables = new JSONObject(new String(json_fis.readAllBytes()));
        json_fis.close();

        XSSFWorkbook workbook = new XSSFWorkbookFactory().create(input_fis);
        Iterator<Sheet> sheets = workbook.sheetIterator();

        while(sheets.hasNext()) {
            Sheet sheet = sheets.next();
            System.out.println(sheet.getSheetName());
            int row_index = -1;
            while(true) {
                row_index += 1;

                Row row = sheet.getRow(row_index);
                if (row == null) break;

                Iterator<Cell> cells = row.cellIterator();
                while(cells.hasNext()) {
                    Cell cell = cells.next();

                    switch(cell.getCellType()) {
                        case STRING:
                            String value = cell.getStringCellValue();
                            Matcher m = pattern.matcher(value);
                            if (m.find()) {
                                String variable = m.group(1);
                                String query = "/" + variable.replaceAll("\\.", "/");
                                cell.setCellValue(dig(variables, query));
                            }
                            break;
                        default:
                            break;
                    }
                }

            }
        }

        FileOutputStream fos = new FileOutputStream(output);
        workbook.write(fos);
    }

    public static String dig(JSONObject variables, String query) {
        return variables.query(query).toString();
    }
}
