package org.example.util;

import java.sql.Date;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;

public class DateUtil {
    public static LocalDateTime getLocalDateTime(String dateTimeStr){
        return LocalDateTime.parse(dateTimeStr.replace(' ', 'T'),
                DateTimeFormatter.ISO_DATE_TIME);
    }

    public static String getString(LocalDateTime localDateTime){
        return localDateTime.format(DateTimeFormatter.ISO_DATE_TIME);
    }
}
