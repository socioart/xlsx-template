package com.socioart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import java.io.*;
import java.util.Iterator;
import java.util.regex.Pattern;
import java.util.regex.Matcher;

/**
 * Hello world!
 *
 */
public class XlsxTemplate
{
    public static void main( String[] args ) throws IOException
    {
        FileInputStream fis;
        String input = args[0];
        String output = args[1];
        // String json = args[2];
        Pattern pattern = Pattern.compile("\\{\\{(.*)\\}\\}");

        try {
            fis = new FileInputStream(input);
        } catch(FileNotFoundException e) {
            System.err.println( "File not found" );
            System.exit(1);
            return;
        }

        XSSFWorkbook workbook = new XSSFWorkbookFactory().create(fis);
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
                                String[] parts = variable.split("\\.");
                                cell.setCellValue(String.join("->", parts));
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
}
