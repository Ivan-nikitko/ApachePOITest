package org.example;

import com.nhl.dflib.DataFrame;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


public class Main {
    private static final String TEMPLATE_PATH = "TemplateBook.xlsx";
    private static final String EXPORT_PATH = "OutTestBook.xlsx";

    public static void main(String[] args) {

        ExcelExporter excelExporter = new ExcelExporter();
        DataFrame artists = createArtists();

        try (FileInputStream inputStream = new FileInputStream(TEMPLATE_PATH) ){
            try (FileOutputStream outputStream = new FileOutputStream(EXPORT_PATH)){
              excelExporter.export(inputStream,outputStream,"Artist", artists);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static DataFrame createArtists() {
        return DataFrame
                .byArrayRow("ID", "ARTIST_NAME", "YEAR_OF_BIRTH")
                .appender()
                .append(1, "Pablo Picasso", 1881)
                .append(2, "Marc Chagall", 1887)
                .append(3, "Salvador Dali", 1904)
                .append(4, "Andy Warhol", 1928)
                .toDataFrame();
    }
}
