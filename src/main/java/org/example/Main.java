package org.example;

import com.nhl.dflib.DataFrame;


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
        excelExporter.export("WriteTestBook.xlsx",  "Artist", createArtists());
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
