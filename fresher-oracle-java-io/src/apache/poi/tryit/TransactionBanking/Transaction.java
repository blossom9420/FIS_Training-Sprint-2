package apache.poi.tryit.TransactionBanking;

import java.time.LocalDateTime;

public class Transaction {
    private Integer id;

    private TransactionType type;

    private BankAccount bankAccount;
    private Double amount;
    private String message;
    private LocalDateTime dateTime;


    public Transaction(Integer id, TransactionType type, BankAccount bankAccount, Double amount, String message) {
        this.id = id;
        this.type = type;
        this.bankAccount = bankAccount;
        this.amount = amount;
        this.message = message;
        this.dateTime = LocalDateTime.now();
    }

    public Transaction() {
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public TransactionType getType() {
        return type;
    }

    public void setType(TransactionType type) {
        this.type = type;
    }

    public BankAccount getBankAccount() {
        return bankAccount;
    }

    public void setBankAccount(BankAccount bankAccount) {
        this.bankAccount = bankAccount;
    }

    public Double getAmount() {
        return amount;
    }

    public void setAmount(Double amount) {
        this.amount = amount;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public LocalDateTime getDateTime() {
        return dateTime;
    }

    public void setDateTime(LocalDateTime dateTime) {
        this.dateTime = dateTime;
    }

    @Override
    public String toString() {
        return "Transaction{" +
                "id=" + id +
                ", type=" + type +
                ", bankAccount=" + bankAccount +
                ", amount=" + amount +
                ", message='" + message + '\'' +
                ", dateTime=" + dateTime +
                '}';
    }
}
