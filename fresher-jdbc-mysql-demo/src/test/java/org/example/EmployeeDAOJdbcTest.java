package org.example;

import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertTrue;

public class EmployeeDAOJdbcTest {
    @Test
    public void testGetAll() {
        IEmployeeDAO employeeDAO = new EmployeeDAOJdbc();
        List<Employee> employeeList = employeeDAO.getAll();

        System.out.println(employeeList);
    }

    @Test
    void add() {
        IEmployeeDAO employeeDAO = new EmployeeDAOJdbc();
        Employee employee = new Employee(20, "Nguyen Thanh Nhan", 103.);

        assertTrue(employeeDAO.add(employee));

        //System.out.println(employeeDAO.getAll());
    }

    @Test
    public void testAddAll() {
        IEmployeeDAO employeeDAO = new EmployeeDAOJdbc();
        List<Employee> list = new ArrayList<>();
        list.add( new Employee(10, "Nguyen Thanh Nhan", 103.) );
        list.add( new Employee(11, "Nguyen Thanh Nhan", 103.) );
        list.add( new Employee(12, "Nguyen Thanh Nhan", 103.) );
        employeeDAO.addAll(list);
        System.out.println(employeeDAO.getAll());
    }
}