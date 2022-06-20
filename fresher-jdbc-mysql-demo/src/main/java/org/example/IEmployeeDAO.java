package org.example;

import java.util.List;

public interface IEmployeeDAO {
    public List<Employee> getAll();

    public boolean add(Employee employee);

    public boolean addAll(List<Employee> employee);
}
