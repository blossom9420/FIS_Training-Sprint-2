package apache.poi.tryit.TransactionBanking;

public class BankAccount {
    private Integer id;
    private String username;

    public BankAccount(Integer id, String username) {
        this.id = id;
        this.username = username;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    @Override
    public String toString() {
        return "BankAccount{" +
                "id=" + id +
                ", username='" + username + '\'' +
                '}';
    }
}
