package org.example;

import lombok.Data;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@Data
public class Employee {
    private long id;
    private String name;
    private double salary;
    private LocalDateTime createDate;

    public Employee(long id, String name, double salary) {
        this.id = id;
        this.name = name;
        this.salary = salary;
        this.createDate = LocalDateTime.now();
    }

    @Override
    public String toString() {
        return "Employee{" +
                "id=" + id +
                ", name='" + name + '\'' +
                ", salary=" + salary +
                ", createDate=" + createDate.format(DateTimeFormatter.ISO_LOCAL_DATE_TIME) +
                '}' +"\n";
    }
}
