package apache.poi.tryit.TransactionBanking;

import apache.poi.tryit.DemoBookExcel.Book;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ReadTransaction {
    public static final int COLUMN_INDEX_ID = 0;
    public static final int COLUMN_INDEX_TRANSACTIONTYPE = 1;
    public static final int COLUMN_INDEX_BANKACCOUNT = 2;
    public static final int COLUMN_INDEX_AMOUNT = 3;
    public static final int COLUMN_INDEX_MESSAGE = 4;
    public static final int COLUMN_INDEX_DATETIME = 5;

    public static void main(String[] args) throws IOException {
        final String excelFilePath = "transaction.xlsx";
        final List<Transaction> transactions = readExcel(excelFilePath);
        for (Transaction transaction : transactions) {
            System.out.println(transaction);
        }
    }

    public static List<Transaction> readExcel(String excelFilePath) throws IOException {
        List<Transaction> listTransactions = new ArrayList<>();

        // Get file
        InputStream inputStream = new FileInputStream(new File(excelFilePath));

        // Get workbook
        Workbook workbook = getWorkbook(inputStream, excelFilePath);

        // Get sheet
        Sheet sheet = workbook.getSheetAt(0);

        // Get all rows
        Iterator<Row> iterator = sheet.iterator();
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            if (nextRow.getRowNum() == 0) {
                // Ignore header
                continue;
            }

            // Get all cells
            Iterator<Cell> cellIterator = nextRow.cellIterator();

            // Read cells and set value for book object
            Transaction transaction = new Transaction();
            while (cellIterator.hasNext()) {
                //Read cell
                Cell cell = cellIterator.next();
                Object cellValue = getCellValue(cell);
                if (cellValue == null || cellValue.toString().isEmpty()) {
                    continue;
                }
                // Set value for book object
                int columnIndex = cell.getColumnIndex();
                switch (columnIndex) {
                    case COLUMN_INDEX_ID:
                        transaction.setId(new BigDecimal((double) cellValue).intValue());
                        break;
                    case COLUMN_INDEX_TRANSACTIONTYPE:
                        transaction.setType(typeByCellValue(getCellValue(cell).toString()));
                        break;
                    case COLUMN_INDEX_BANKACCOUNT:
                        transaction.setBankAccount(bankAccounByCellValue((new BigDecimal((double) cellValue).intValue())));
                        break;
                    case COLUMN_INDEX_AMOUNT:
                        transaction.setAmount((Double) getCellValue(cell));
                        break;
                    case COLUMN_INDEX_MESSAGE:
                        transaction.setMessage((String) getCellValue(cell));
                        break;
                    case COLUMN_INDEX_DATETIME:
                        transaction.setDateTime(LocalDateTime.parse((getCellValue(cell)).toString()));
                        break;
                    default:
                        break;
                }

            }
            listTransactions.add(transaction);
        }

        workbook.close();
        inputStream.close();

        return listTransactions;
    }

    private static TransactionType typeByCellValue(String type) {
        if (type.equals("DEPOSIT"))
            return TransactionType.DEPOSIT;
        if (type.equals("TRANSFER"))
            return TransactionType.TRANSFER;
        return TransactionType.WITHDRAW;
    }

    private static BankAccount bankAccounByCellValue(int id) {
        if (id == 1)
            return new BankAccount(1, "FIS Training");
        return new BankAccount(2, "Phạm Hoàng Long");
    }


    // Get Workbook
    private static Workbook getWorkbook(InputStream inputStream, String excelFilePath) throws IOException {
        Workbook workbook = null;
        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        return workbook;
    }

    // Get cell value
    private static Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellTypeEnum();
        Object cellValue = null;
        switch (cellType) {
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case FORMULA:
                Workbook workbook = cell.getSheet().getWorkbook();
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                cellValue = evaluator.evaluate(cell).getNumberValue();
                break;
            case NUMERIC:
                cellValue = cell.getNumericCellValue();
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case _NONE:
            case BLANK:
            case ERROR:
                break;
            default:
                break;
        }

        return cellValue;
    }
}
