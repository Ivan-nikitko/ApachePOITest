package org.example;

import org.example.model.Artist;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;



public class Main {

    public static void main(String[] args) {
        ExcelExporter excelExporter = new ExcelExporter();
        excelExporter.export("WriteTestBook.xlsx","19th century artists","Artist19", create19ThCenturyArtistList());
        excelExporter.export("WriteTestBook.xlsx","20th century artists","Artist20", create20ThCenturyArtistList());
    }

    private static List<Artist> create19ThCenturyArtistList() {
        Artist picasso = new Artist(1, "Pablo Picasso", "25 October 1881");
        Artist chagall = new Artist(2, "Marc Chagall", "24 June 1887");
        return new ArrayList<>(Arrays.asList(picasso, chagall));
    }

    private static List<Artist> create20ThCenturyArtistList() {
        Artist picasso = new Artist(3, "Salvador Dali", "11 May 1904");
        Artist chagall = new Artist(4, "Andy Warhol", "6 August 1928");
        return new ArrayList<>(Arrays.asList(picasso, chagall));
    }
}
