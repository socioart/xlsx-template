package com.socioart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import java.io.*;
import java.util.Iterator;
import java.util.ArrayList;
import java.util.regex.Pattern;
import java.util.regex.Matcher;
import org.json.*;

public class XlsxTemplate
{
    static class Template {
        static class Directive {}
        static class IfDirective extends Directive {
            public String variable;
            IfDirective(String _variable) {
                variable = _variable;
            }
        }
        static class EachDirective extends Directive {
            public String variable;
            EachDirective(String _variable) {
                variable = _variable;
            }
        }

        private Pattern _pattern_variable;
        private Pattern _pattern_if;
        private Pattern _pattern_each;
        private XSSFWorkbook _template;
        private JSONObject _variables;

        public Template() {
            _pattern_variable = Pattern.compile("\\{\\{(.*)\\}\\}");
            _pattern_if = Pattern.compile("\\{\\{#if (.*)\\}\\}");
            _pattern_each = Pattern.compile("\\{\\{#each (.*)\\}\\}");
        }

        public void render(XSSFWorkbook template, JSONObject variables, String output) throws IOException {
            _template = template;
            _variables = variables;

            Iterator<Sheet> sheets = _template.sheetIterator();

            while(sheets.hasNext()) {
                Sheet sheet = sheets.next();

                // A1 が "beforeRow" でなければ無視
                Row firstRow = sheet.getRow(0);
                if (firstRow == null) continue;

                Cell firstCell = firstRow.getCell(0);
                if (firstCell == null) continue;
                if (firstCell.getCellType() != CellType.STRING || !firstCell.getStringCellValue().equals("beforeRow")) continue;

                int row_index = 0;
                while(true) {
                    if (sheet.getLastRowNum() < row_index) break;

                    Row row = sheet.getRow(row_index);
                    row_index += processTemplateRow(sheet, row) + 1;
                }

                sheet.setColumnHidden(0, true); // hide meta cells
            }

            FileOutputStream fos = new FileOutputStream(output);
            _template.write(fos);
        }

        // return number of inserted row (maybe negative)
        private int processTemplateRow(Sheet sheet, Row row) {
            if (row == null) return 0;

            Cell metaCell = row.getCell(0);
            Directive directive = null;

            if (metaCell != null && metaCell.getCellType() == CellType.STRING) {
                directive = parseDirective(metaCell.getStringCellValue());
                metaCell.setBlank();
            }

            if (directive instanceof IfDirective) {
                return processIfDirectiveRow(sheet, row, (IfDirective)directive);
            } else if (directive instanceof EachDirective) {
                return processEachDirectiveRow(sheet, row, (EachDirective)directive);
            } else {
                return processNoDirectiveRow(sheet, row);
            }
        }

        private int processNoDirectiveRow(Sheet sheet, Row row) {
            replaceCells(row, _variables);
            return 0;
        }

        private int processIfDirectiveRow(Sheet sheet, Row row, IfDirective directive) {
            if (!isTruthy(((IfDirective)directive).variable, _variables)) {
                removeRow(sheet, row);
                return -1; // 削除したので -1
            } else {
                return processNoDirectiveRow(sheet, row);
            }
        }

        private int processEachDirectiveRow(Sheet sheet, Row template_row, EachDirective directive) {
            int template_row_index = template_row.getRowNum();

            ArrayList<CellRangeAddress> merged_regions_in_row = new ArrayList<CellRangeAddress>();
            for(CellRangeAddress mr: sheet.getMergedRegions()) {
                if (mr.getFirstRow() == template_row_index && mr.getLastRow() == template_row_index) merged_regions_in_row.add(mr);
            }

            JSONArray items = digArray(((EachDirective)directive).variable, _variables);
            int row_index = template_row_index;

            for(Object obj: items) {
                JSONObject item = (JSONObject)obj;
                row_index += 1;

                sheet.shiftRows(row_index, sheet.getLastRowNum(), 1, true, true);
                Row new_row = sheet.createRow(row_index);
                copyRow(template_row, new_row);

                // Copy merged regions
                for(CellRangeAddress mr: merged_regions_in_row) {
                    sheet.addMergedRegion(
                        new CellRangeAddress(row_index, row_index, mr.getFirstColumn(), mr.getLastColumn())
                    );
                }

                replaceCells(new_row, item);
            }

            removeRow(sheet, template_row);
            return items.length() - 1; // アイテムの数 - テンプレートの行
        }

        // https://stackoverflow.com/questions/5785724/how-to-insert-a-row-between-two-rows-in-an-existing-excel-with-hssf-apache-poi
        private void copyRow(Row src, Row dest) {
            for (int i = 0; i < src.getLastCellNum(); i++) {
                Cell src_cell = src.getCell(i);
                if (src_cell == null) continue;

                Cell dest_cell = dest.createCell(i);

                dest_cell.setCellStyle(src_cell.getCellStyle());
                if (src_cell.getCellComment() != null) dest_cell.setCellComment(src_cell.getCellComment());
                if (src_cell.getHyperlink() != null) dest_cell.setHyperlink(src_cell.getHyperlink());

                switch (src_cell.getCellType()) {
                    case _NONE:
                        break;
                    case BLANK:
                        dest_cell.setBlank();
                        break;
                    case BOOLEAN:
                        dest_cell.setCellValue(src_cell.getBooleanCellValue());
                        break;
                    case ERROR:
                        dest_cell.setCellErrorValue(src_cell.getErrorCellValue());
                        break;
                    case FORMULA:
                        dest_cell.setCellFormula(src_cell.getCellFormula());
                        break;
                    case NUMERIC:
                        dest_cell.setCellValue(src_cell.getNumericCellValue());
                        break;
                    case STRING:
                        dest_cell.setCellValue(src_cell.getRichStringCellValue());
                        break;
                }
            }
        }

        private void removeRow(Sheet sheet, Row row) {
            int row_index = row.getRowNum();

            // Update merged cells
            ArrayList<Integer> merged_region_indices_to_delete = new ArrayList<Integer>();
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                CellRangeAddress mr = sheet.getMergedRegion(i);
                if (mr.getFirstRow() == row_index && mr.getLastRow() == row_index) {
                    merged_region_indices_to_delete.add(i);
                }
            }
            sheet.removeMergedRegions(merged_region_indices_to_delete);

            sheet.removeRow(row); // remove row content
            sheet.shiftRows(row_index + 1, sheet.getLastRowNum(), -1);
        }

        private void replaceCells(Row row, JSONObject locals) {
            Iterator<Cell> cells = row.cellIterator();
            if (cells.hasNext()) cells.next(); // skip metaCell
            while(cells.hasNext()) {
                Cell cell = cells.next();

                switch(cell.getCellType()) {
                    case STRING:
                        String value = cell.getStringCellValue();
                        Matcher m = _pattern_variable.matcher(value);
                        if (m.find()) {
                            String query = toJSONPointer(m.group(1));
                            cell.setCellValue(digString(query, locals));
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        // https://handlebarsjs.com/guide/builtin-helpers.html#if
        private boolean isTruthy(String query, JSONObject locals) {
            Object obj = locals.query(query);
            if (obj == null) return false;
            if (obj instanceof Boolean) return (boolean)obj;
            if (obj instanceof String) return !((String)obj).isEmpty();
            if (obj instanceof Integer) return (int)obj != 0;
            if (obj instanceof JSONArray) return !((JSONArray)obj).isEmpty();
            return true;
        }

        private String digString(String query, JSONObject locals) {
            Object val = locals.query(query);
            if (val == null) return null;
            return val.toString();
        }

        private JSONArray digArray(String query, JSONObject locals) {
            Object obj = locals.query(query);
            if (obj instanceof JSONArray) return (JSONArray)obj;
            return null;
        }

        private Directive parseDirective(String v) {
            Matcher m;

            m = _pattern_if.matcher(v);
            if (m.find()) {
                return new IfDirective(toJSONPointer(m.group(1)));
            }

            m = _pattern_each.matcher(v);
            if (m.find()) {
                return new EachDirective(toJSONPointer(m.group(1)));
            }

            return null;
        }

        private String toJSONPointer(String handlebarsVariableRef) {
            return "/" + handlebarsVariableRef.replaceAll("\\.", "/");
        }
    }

    public static void main( String[] args ) throws IOException
    {
        FileInputStream input_fis;
        FileInputStream json_fis;

        String subcommand = args[0];
        String input = args[1];

        try {
            input_fis = new FileInputStream(input);
        } catch(FileNotFoundException e) {
            System.err.println( "Templatel file not found" );
            System.exit(1);
            return;
        }

        XSSFWorkbook workbook = new XSSFWorkbookFactory().create(input_fis);
        input_fis.close();

        if (subcommand.equals("compile")) {
            String output = args[2];
            String json_file = args[3];

            try {
                json_fis = new FileInputStream(json_file);
            } catch(FileNotFoundException e) {
                System.err.println( "Data file not found" );
                System.exit(1);
                return;
            }

            JSONObject variables = new JSONObject(new String(json_fis.readAllBytes()));
            json_fis.close();

            Template template = new Template();
            template.render(workbook, variables, output);
        } else if (subcommand.equals("list-pictures")) {
            listPictures(workbook);
        } else if (subcommand.equals("replace-picture")) {
            String output = args[2];
            String cell_ref = args[3];
            String image_path = args[4];

            replacePicture(workbook, output, cell_ref, image_path);
        }
    }

    static private void replacePicture(Workbook workbook, String output, String cell_ref_string, String image_path) throws IOException {
        CellReference cell_ref = new CellReference(cell_ref_string);
        String sheet_name = cell_ref.getSheetName();

        FileInputStream fis = new java.io.FileInputStream(image_path) ;
        byte[] data = fis.readAllBytes();
        fis.close();
        int new_picture_index = workbook.addPicture(data, Workbook.PICTURE_TYPE_PNG);

        Iterator<Sheet> sheets = workbook.sheetIterator();
        while(sheets.hasNext()) {
            Sheet sheet = sheets.next();
            if (sheet_name != null && !sheet.getSheetName().equals(sheet_name)) continue;

            XSSFDrawing drawing = (XSSFDrawing)sheet.getDrawingPatriarch();
            if (drawing == null) continue;

            for (XSSFShape shape : drawing.getShapes()){
                if (!(shape instanceof XSSFPicture)) continue;
                XSSFPicture picture = (XSSFPicture)shape;

                XSSFClientAnchor xca = picture.getPreferredSize();
                if (xca.getRow1() == cell_ref.getRow() && xca.getCol1() == cell_ref.getCol()) {
                    drawing.createPicture(xca, new_picture_index);
                }
            }
        }

        FileOutputStream fos = new FileOutputStream(output);
        workbook.write(fos);
    }

    static private void listPictures(Workbook workbook) {
        System.out.println("cell,mime_type");

        Iterator<Sheet> sheets = workbook.sheetIterator();
        while(sheets.hasNext()) {
            Sheet sheet = sheets.next();

            XSSFDrawing drawing = (XSSFDrawing)sheet.getDrawingPatriarch();
            if (drawing == null) continue;

            for (XSSFShape shape : drawing.getShapes()){
                if (!(shape instanceof XSSFPicture)) continue;
                XSSFPicture picture = (XSSFPicture)shape;
                XSSFPictureData picture_data = picture.getPictureData();

                // 画像があるセルを取得
                XSSFClientAnchor xca = picture.getPreferredSize();
                CellReference cell_ref = new CellReference(sheet.getSheetName(), xca.getRow1(), xca.getCol1(), false, false);

                // TODO: escape sheet name
                // FIY: picture_data.getPackagePart().getPartName() to get xlsx path
                System.out.println(cell_ref.formatAsString(true) + "," + picture_data.getMimeType());
            }
        }
    }
}
