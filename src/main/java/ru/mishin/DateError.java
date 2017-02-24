package ru.mishin;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

import static java.util.Locale.ENGLISH;

/**
 * Created by ira on 24.02.2017.
 */
public class DateError {
    public static void main(String[] args) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("d,MM,yyyy");
        String dateOfBirth="01,01,1772";
        LocalDate localDate = LocalDate.parse(dateOfBirth, formatter);
        System.out.println(localDate.format(DateTimeFormatter.ofPattern("d MMM yyyy",ENGLISH)));
    }
}
