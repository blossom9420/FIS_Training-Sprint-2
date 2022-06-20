package com.slf4j;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Hello world for slf4J!
 *
 */
public class App {
	public static void main(String[] args) {
		Logger logger = LoggerFactory.getLogger(App.class);
		
		logger.info("This is an information message");
		logger.error("this is a error message");
		logger.warn("this is a warning message");
		logger.debug("this is a warning message");
	}
}
