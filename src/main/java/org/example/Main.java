package org.example;

import org.example.model.Artist;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;


public class Main {

    public static void main(String[] args) {

        ExcelExporter excelExporter = new ExcelExporter();
//        try ( FileInputStream inputStream = new FileInputStream("WriteTestBook.xlsx")){
//
//            XSSFWorkbook artist19 = excelExporter.export1(inputStream, "19th century artists", "Artist19", create19ThCenturyArtistList());
//            XSSFWorkbook artist20 = excelExporter.export1(inputStream, "20th century artists", "Artist20", create20ThCenturyArtistList());
//            FileOutputStream outputStream = new FileOutputStream("WriteTestBook.xlsx");
//            artist19.write(outputStream);
//            artist20.write(outputStream);
//            System.out.println("Data exported");
//        } catch (IOException e) {
//            System.out.println("Error exporting data: " + e.getMessage());
//
//        }

        excelExporter.export("WriteTestBook.xlsx", "Artist", "Artist", createArtistList());
       // excelExporter.export("WriteTestBook.xlsx", "20th century artists", "Artist20", create20ThCenturyArtistList());
    }

    private static List<Artist> createArtistList() {
        Artist picasso = new Artist(1, "Pablo Picasso", "1881");
        Artist chagall = new Artist(2, "Marc Chagall", "1887");
        return new ArrayList<>(Arrays.asList(picasso, chagall));
    }

    private static List<Artist> create20ThCenturyArtistList() {
        Artist dali = new Artist(3, "Salvador Dali", "11 May 1904");
        Artist warhol = new Artist(4, "Andy Warhol", "6 August 1928");
        return new ArrayList<>(Arrays.asList(dali, warhol));
    }
}
