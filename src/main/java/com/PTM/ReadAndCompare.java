package com.PTM;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;

import java.io.*;
import java.util.List;
import java.util.Scanner;

public class ReadAndCompare {

    //Read table cell
//    @Test
    public void readDocxTable(String wordFile, String prtFile) throws IOException, InvalidFormatException {

        FileInputStream fis = new FileInputStream(wordFile);
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(fis));
        List<XWPFTable> tables = doc.getTables();
        for (XWPFTable table : tables) {

            if (table.getRow(0).getCell(0).getText().contains("Tag")) {

                for (XWPFTableRow row : table.getRows()) {
                    if (row.getCell(0).getText().contains("Tag") && row.getCell(1).getText().contains("Sub 1")
                            && row.getCell(2).getText().contains("Sub 2") && row.getCell(3).getText().contains("Sub 3")
                            && row.getCell(4).getText().contains("Description")) {
                        continue;
                    }
                    if (!row.getCell(0).getText().startsWith("0")) {
                        continue;
                    }
                    if (row.getCell(0).getText().contains("n")) {
                        continue;
                    }
                    String cellOne = row.getCell(0).getText();
                    String cellTwo = row.getCell(1).getText();
                    String cellThree = row.getCell(2).getText();
                    String cellFour = row.getCell(3).getText();
                    String searchString = cellOne + cellTwo + cellThree + cellFour;
//                    System.out.println(searchString);
                    int wordCount = searchString(prtFile, searchString);
                    System.out.println(searchString + " appears " + wordCount + " times");

                }
            }

        }

    }


    private int searchString(String file, String searchString) throws IOException {

        BufferedReader bf = new BufferedReader(new FileReader("TC602_C1_AMD_196090.prt"));
        int linecount = 0;
        String line;

        System.out.println("Searching for " + searchString);

        while ((line = bf.readLine()) != null) {
            if (line.contains(searchString))
                linecount++;
        }
        bf.close();
        return linecount;
    }
}



