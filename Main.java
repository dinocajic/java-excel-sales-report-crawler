import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        List<Invoice> invoices = new ArrayList<>();
        String branch   = "Branch";
        String date     = "02-05-2020";
        String filePath = "pathtofile";

        try( Workbook wb = WorkbookFactory.create(
                new File(filePath) )
        ) {
            // Create Sheet
            Sheet sheet = wb.getSheetAt(0);

            // Iterate through rows
            int rowStart = 0;
            int rowEnd   = sheet.getLastRowNum();

            for( int i = rowStart; i < rowEnd; i++ ) {
                Row row = sheet.getRow(i);

                if ( row.getCell(0) == null && row.getCell(7) == null )
                    continue;

                if ( row.getCell(0) != null && row.getCell(0).getStringCellValue().startsWith("200") ) {
                    // We need to create a new Invoice and add the items to it
                    Invoice invoice = new Invoice();
                    Invoice.number_of_items++; // Needed for excel generation

                    invoice.setInvoice_number( row.getCell(0).getStringCellValue() );
                    invoice.setType( row.getCell(4).getStringCellValue() );
                    invoice.setAccount( row.getCell(5).getStringCellValue() );
                    invoice.setName( row.getCell(8).getStringCellValue() );
                    invoice.setJob_num( row.getCell(14).getNumericCellValue() + "" );
                    invoice.setSales_user( row.getCell(17).getStringCellValue() );
                    invoice.setSis_mat_br( (int) row.getCell(20).getNumericCellValue() + "" );
                    invoice.setDate( row.getCell(25).getStringCellValue() );
                    invoice.setTotal( row.getCell(26).getNumericCellValue() + "" );
                    invoice.setCost( row.getCell(29).getNumericCellValue() + "" );
                    invoice.setGm_perc( "0" );

                    List<String[]> items = new ArrayList<>();

                    // Now we need to get the items sold
                    for ( int j = i + 2; j <= rowEnd; j++ ) {
                        row = sheet.getRow(j);

                        if ( isRowEmpty( row ) )
                            continue;

                        if ( row.getCell(0) != null ) {
                            i = j - 1; // Have to subtract 1 since outer loop increments i
                            break;
                        }

                        // Can skip multiple line descriptions and just focus on the first line
                        if ( row.getCell(2) == null )
                            continue;

                        if ( getStringValue(row, 2).contains("Item") )
                            continue;

                        String[] item = {
                            getStringValue( row, 2 ),                                  // Item Number - Can be String or Numeric
                            getStringValue( row, 7 ),                                  // Description
                            getStringValue( row, 13 ),                                 // Quantity
                            row.getCell(15).getStringCellValue(),                           // U/M
                            String.format( "%.2f", row.getCell(19).getNumericCellValue() ), // Unit Price
                            row.getCell(20).getStringCellValue(),                           // U/M
                            String.format( "%.2f", row.getCell(24).getNumericCellValue() ), // Unit Cost
                            String.format( "%.2f", row.getCell(27).getNumericCellValue() )  // GM%
                        };

                        // Check to see if item number has been split into two rows
                        // Using try/catch since we might be on the last line but we're trying to catch a
                        // row that does not exist
                        try {
                            row = sheet.getRow( j + 1 );

                            if ( row.getCell(2) != null && row.getCell(13) == null ) {
                                item[0] = item[0] + getStringValue( row, 2 );
                                j++; // Since we grabbed the next row, we need to skip over it
                            }
                        } catch (NullPointerException ignored) { }

                        items.add( item );
                        Invoice.number_of_items++; // Needed for excel generation
                    }

                    invoice.setItems( items );

                    // add invoice to list of invoices
                    invoices.add( invoice );
                }
            }

        } catch( Exception e ) {
            e.printStackTrace();
        }

        // For quick reporting validation
        for ( Invoice invoice : invoices ) {
            System.out.println(invoice);
        }

        // Create Excel sheet
        generateExcelFile( invoices, branch, date );
    }

    /**
     * Returns a String representation of the cell since a cell could be numeric or a string
     * @param row
     * @param column
     * @return
     */
    private static String getStringValue( Row row, int column ) {
        String value;

        if (row.getCell( column ).getCellType().equals( CellType.STRING )) {
            value = row.getCell( column ).getStringCellValue();
        } else {
            Object item = row.getCell( column ).getNumericCellValue();
            value = new BigDecimal( item.toString() ).toPlainString();
        }

        return value;
    }

    /**
     * Checks to see if the entire row is empty
     * @param row
     * @return
     */
    public static boolean isRowEmpty(Row row) {
        boolean isEmpty = true;
        DataFormatter dataFormatter = new DataFormatter();

        if( row != null ) {
            for(Cell cell: row) {
                if(dataFormatter.formatCellValue(cell).trim().length() > 0) {
                    isEmpty = false;
                    break;
                }
            }
        }

        return isEmpty;
    }

    /**
     * Creates the excel file from the generated data
     * @param invoices
     * @param branch
     * @param date
     */
    private static void generateExcelFile( List<Invoice> invoices, String branch, String date ) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet       = workbook.createSheet("Report");
        sheet.getFooter().setCenter( "Page: &P of &N" );

        // Total number of items plus some extra content for each invoice, which takes up 4 rows, plus
        // at least 4 more rows for header. Just to be on the safe side, I added 7.
        Object[][] excelData = new Object[ Invoice.number_of_items + (invoices.size() * 4) + 7][11];

        int i = 0;

        excelData[i++] = new Object[]{ "Sales Report: " + branch };
        excelData[i++] = new Object[]{ "Date: " + invoices.get(0).getDate() };
        excelData[i++] = new Object[]{ "" };

        excelData[i++] = new Object[]{ "Invoice #", "Type", "Account", "Name", "Sales #", "Branch",
                "Date", "Total", "Cost", "Total GM%"
        };

        excelData[i++] = new Object[]{ "" };

        sheet.setRepeatingRows(CellRangeAddress.valueOf("1:5"));
        sheet.addMergedRegion( new CellRangeAddress(0,0,0,9) );

        for ( Invoice invoice : invoices ) {
            excelData[i++] = new Object[]{
                    invoice.getInvoice_number(),
                    invoice.getType(),
                    invoice.getAccount(),
                    invoice.getName(),
                    invoice.getSales_user(),
                    invoice.getSis_mat_br(),
                    invoice.getDate(),
                    invoice.getTotal(),
                    invoice.getCost(),
                    invoice.getGm_perc()
            };

            excelData[i++] = new Object[] {""};

            excelData[i++] = new Object[] { "", "", "Item Number", "Description", "Qty",
                "U/M", "Unit Price", "U/M", "Unit Cost", "GM%"
            };

            for ( String[] item : invoice.getItems() ) {
                excelData[i++] = new Object[] {
                        "",
                        "",
                        item[0], // Item Number
                        item[1], // Description
                        item[2], // Qty
                        item[3], // U/M
                        item[4], // Unit Price
                        item[5], // U/M
                        item[6], // Unit Cost
                        item[7]  // GM%
                };
            }

            excelData[i++] = new Object[] {""};
        }

        excelData[i++] = new Object[] { "", "", "", "", "", "",
                "Grand Total",
                String.format("%.2f", Invoice.static_total),
                String.format("%.2f", Invoice.static_cost),
                String.format( "%.2f", ( ( Invoice.static_total - Invoice.static_cost ) / Invoice.static_total * 100 ) )
        };

        excelData[i] = new Object[] { "", "", "", "", "", "",
                "Additional Charges",
                String.format("%.2f", Invoice.static_additional_charges),
                "¯\\_(ツ)_/¯" // Just means that I don't know what the cost is
        };

        // Styles
        XSSFCellStyle top_heading = workbook.createCellStyle();
        Font top_font = workbook.createFont();
        top_font.setBold( true );
        top_font.setFontHeightInPoints( (short) 20 );
        top_heading.setFont( top_font );

        XSSFCellStyle main_heading = workbook.createCellStyle();
        Font bold = workbook.createFont();
        bold.setBold(true);
        main_heading.setFont(bold);
        main_heading.setBorderBottom( BorderStyle.MEDIUM );

        XSSFCellStyle item_heading = workbook.createCellStyle();
        item_heading.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        item_heading.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font italic = workbook.createFont();
        italic.setItalic(true);
        item_heading.setFont( italic );

        List<String> main_headings = new ArrayList<>();
        main_headings.add("Invoice #");
        main_headings.add("Account");
        main_headings.add("Type");
        main_headings.add("Name");
        main_headings.add("Sales #");
        main_headings.add("Branch");
        main_headings.add("Date");
        main_headings.add("Total");
        main_headings.add("Cost");
        main_headings.add("Total GM%");

        List<String> item_headings = new ArrayList<>();
        item_headings.add("Item Number");
        item_headings.add("Description");
        item_headings.add("Qty");
        item_headings.add("U/M");
        item_headings.add("Unit Price");
        item_headings.add("U/M");
        item_headings.add("Unit Cost");
        item_headings.add("GM%");

        int rowCount = 0;

        for (Object[] excelRow : excelData) {
            Row row = sheet.createRow( rowCount++ );

            int columnCount = 0;

            for (Object field : excelRow) {
                Cell cell = row.createCell( columnCount++ );

                if ( field == null )
                    break;

                if ( rowCount == 1 )
                    cell.setCellStyle( top_heading );

                if ( main_headings.contains( field.toString() ) )
                    cell.setCellStyle( main_heading );

                if ( item_headings.contains( field.toString() ) )
                    cell.setCellStyle( item_heading );

                if (field instanceof String) {
                    cell.setCellValue( (String) field );
                } else if (field instanceof Integer) {
                    cell.setCellValue( (Integer) field );
                }
            }
        }

        try (FileOutputStream outputStream = new FileOutputStream(
                "SalesReport_" +
                        branch.replace(" ", "_") + "_" +
                        date + "_"  +
                        System.currentTimeMillis() +
                        ".xlsx")
        ) {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
