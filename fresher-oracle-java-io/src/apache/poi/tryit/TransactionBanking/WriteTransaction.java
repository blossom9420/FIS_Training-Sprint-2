package apache.poi.tryit.TransactionBanking;

import apache.poi.tryit.DemoBookExcel.Book;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

public class WriteTransaction {
    public static final int COLUMN_INDEX_ID = 0;
    public static final int COLUMN_INDEX_TRANSACTIONTYPE = 1;
    public static final int COLUMN_INDEX_BANKACCOUNT = 2;
    public static final int COLUMN_INDEX_AMOUNT = 3;
    public static final int COLUMN_INDEX_MESSAGE = 4;
    public static final int COLUMN_INDEX_DATETIME = 5;
    private static CellStyle cellStyleFormatNumber = null;

    public static void main(String[] args) throws IOException {
        final List<Transaction> transactions = getTransactions();
        final String excelFilePath = "C:/sA/transaction.xlsx";
        try{
            writeExcel(transactions, excelFilePath);
        }catch (Exception e){
            System.out.println(" excelFilePath = \"C:/AA/transaction.xlsx\" . Cẩn thận bị sai chỗ lưu file");
        }

    }

    public static void writeExcel(List<Transaction> transactions, String excelFilePath) throws IOException {
        // Create Workbook
        Workbook workbook = getWorkbook(excelFilePath);

        // Create sheet
        Sheet sheet = workbook.createSheet("Transaction"); // Create sheet with sheet name

        int rowIndex = 0;

        // Write header
        writeHeader(sheet, rowIndex);

        // Write data
        rowIndex++;
        for (Transaction transaction : transactions) {
            // Create row
            Row row = sheet.createRow(rowIndex);
            // Write data on row
            writeTransaction(transaction, row);
            rowIndex++;
        }

        // Auto resize column witdth
        int numberOfColumn = sheet.getRow(0).getPhysicalNumberOfCells();
        autosizeColumn(sheet, numberOfColumn);

        // Create file excel
        createOutputFile(workbook, excelFilePath);
        System.out.println("Done!!!");
    }

    // Create dummy data
    private static List<Transaction> getTransactions() {
        List<Transaction> listTransaction = new ArrayList<>();
        Transaction transaction;
        TransactionType type;
        for (int i = 1; i <= 10; i++) {
            type = randomType();
            transaction = new Transaction(i, type, createBankAccountByType(type), (double) i * 1000000,
                    createMessageByType(type));
            listTransaction.add(transaction);
        }
        return listTransaction;
    }

    private static TransactionType randomType() {
        double random = (Math.random()) * 3;
        int num = (int) random + 1;  // 0-2 + 1 = 1-3
        if (num == 1)
            return TransactionType.DEPOSIT;
        if (num == 2)
            return TransactionType.WITHDRAW;
        return TransactionType.TRANSFER;
    }

    private static BankAccount createBankAccountByType(TransactionType type) {
        if (type == TransactionType.WITHDRAW) {
            return new BankAccount(1, "FIS Training");
        } // TRAMSFER or DEPOSIT = Student
        return new BankAccount(2, "Phạm Hoàng Long");
    }

    private static String createMessageByType(TransactionType type) {
        if (type == TransactionType.WITHDRAW) {
            return "FIS Training rút tiền";
        } // TRAMSFER or DEPOSIT = Student
        if (type == TransactionType.DEPOSIT) {
            return "PHL nộp tiền";
        }
        return "1811063643_Phạm Hoàng Long";
    }

    // Create workbook
    private static Workbook getWorkbook(String excelFilePath) throws IOException {
        Workbook workbook = null;

        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook();
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook();
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        return workbook;
    }

    // Write header with format
    private static void writeHeader(Sheet sheet, int rowIndex) {
        // create CellStyle
        CellStyle cellStyle = createStyleForHeader(sheet);

        // Create row
        Row row = sheet.createRow(rowIndex);

        // Create cells
        Cell cell = row.createCell(COLUMN_INDEX_ID);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Id");

        cell = row.createCell(COLUMN_INDEX_TRANSACTIONTYPE);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Transaction Type");

        cell = row.createCell(COLUMN_INDEX_BANKACCOUNT);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Bank Account");

        cell = row.createCell(COLUMN_INDEX_AMOUNT);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Amount");

        cell = row.createCell(COLUMN_INDEX_MESSAGE);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Message");

        cell = row.createCell(COLUMN_INDEX_DATETIME);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("DateTime");


    }

    // Write data
    private static void writeTransaction(Transaction transaction, Row row) {
        if (cellStyleFormatNumber == null) {
            // Format number
            short format = (short) BuiltinFormats.getBuiltinFormat("#,##0");
            // DataFormat df = workbook.createDataFormat();
            // short format = df.getFormat("#,##0");

            //Create CellStyle
            Workbook workbook = row.getSheet().getWorkbook();
            cellStyleFormatNumber = workbook.createCellStyle();
            cellStyleFormatNumber.setDataFormat(format);
        }

        Cell cell = row.createCell(COLUMN_INDEX_ID);
        cell.setCellValue(transaction.getId());

        cell = row.createCell(COLUMN_INDEX_TRANSACTIONTYPE);
        cell.setCellValue(transaction.getType().toString());

        cell = row.createCell(COLUMN_INDEX_BANKACCOUNT);
        cell.setCellValue(transaction.getBankAccount().getId());

        cell = row.createCell(COLUMN_INDEX_AMOUNT);
        cell.setCellValue(transaction.getAmount());
        cell.setCellStyle(cellStyleFormatNumber);


        cell = row.createCell(COLUMN_INDEX_MESSAGE);
        cell.setCellValue(transaction.getMessage());

        cell = row.createCell(COLUMN_INDEX_DATETIME);
        cell.setCellValue(transaction.getDateTime().toString());
    }

    // Create CellStyle for header
    private static CellStyle createStyleForHeader(Sheet sheet) {
        // Create font
        Font font = sheet.getWorkbook().createFont();
        font.setFontName("Times New Roman");
        font.setBold(true);
        font.setFontHeightInPoints((short) 14); // font size
        font.setColor(IndexedColors.WHITE.getIndex()); // text color

        // Create CellStyle
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        return cellStyle;
    }


    // Auto resize column width
    private static void autosizeColumn(Sheet sheet, int lastColumn) {
        for (int columnIndex = 0; columnIndex < lastColumn; columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
        }
    }

    // Create output file
    private static void createOutputFile(Workbook workbook, String excelFilePath) throws IOException {
        try (OutputStream os = new FileOutputStream(excelFilePath)) {
            workbook.write(os);
        }
    }
}
