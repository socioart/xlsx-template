package com.socioart;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import java.io.*;

/**
 * Hello world!
 *
 */
public class XlsxTemplate
{
    public static void main( String[] args ) throws IOException
    {
        try {
            FileInputStream fis = new FileInputStream(args[0]);
            XSSFWorkbook wb = new XSSFWorkbookFactory().create(fis);
        } catch(FileNotFoundException e) {
            System.err.println( "File not found" );
            System.exit(1);
        }
        System.out.println( "Hello World!" );
    }
}
