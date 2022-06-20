package com.slf4j;

import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class PlaceHolders {

    @Test
    public void testParameterizedLogging_01()
    {
        //Creating the Logger object
        Logger logger = LoggerFactory.getLogger(PlaceHolders.class);
        Integer age = 23;

        //Logging the information
        logger.info("At the age of {} ramu got his first job", age);
    }

    @Test
    public void testParameterizedLogging_02()
    {
        Integer oldWeight;
        Integer newWeight;

        oldWeight = 5;
        newWeight = 6;

        //Creating the Logger object
        Logger logger = LoggerFactory.getLogger(PlaceHolders.class);

        //Logging the information
        logger.info("Old weight is {}. new weight is {}.", oldWeight, newWeight);

        //Logging the information
        logger.info("After the program weight reduced is: "+(oldWeight-newWeight));
    }

    @Test
    public void testParameterizedLogging_03()
    {
        Integer age = 24;
        String designation = "Software Engineer";
        String company = "Infosys";

        //Creating the Logger object
        Logger logger = LoggerFactory.getLogger(PlaceHolders.class);

        //Logging the information
        logger.info("At the age of {} ramu got his first job as a {} at {}", age, designation, company);
    }
}
