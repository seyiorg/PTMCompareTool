package com.PTM;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;
import java.io.*;
import java.util.List;

public class ReadWord_copy {

    //Read table cell
    @Test
    public void readDocxTable() throws IOException, InvalidFormatException {

        FileInputStream fis = new FileInputStream("TC602 Vol 22.docx");
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(fis));
        List<XWPFTable> tables = doc.getTables();
        for (XWPFTable table : tables) {
            if (table.getRow(0).getCell(0).getText().contains("Tag")) {

                for ( XWPFTableRow row : table.getRows() ){
                    if (row.getCell(0).getText().contains("Tag") && row.getCell(1).getText().contains("Sub 1")
                            && row.getCell(2).getText().contains("Sub 2") && row.getCell(3).getText().contains("Sub 3")
                            && row.getCell(4).getText().contains("Description")) {
                        continue;
                    }
                    if (!row.getCell(0).getText().startsWith("0")){
                        continue;
                    }
                    String cellOne = row.getCell(0).getText();
                    String cellTwo = row.getCell(1).getText();
                    String cellThree = row.getCell(2).getText();
                    String cellFour = row.getCell(3).getText();
                    String searchString = cellOne + cellTwo + cellThree + cellFour;
                    System.out.println("Concatinated string is: " + searchString);
                    searchString(searchString);

                }
            }

        }

    }


    private void searchString(String searchString) throws IOException{

        //Open File
        BufferedReader bf = new BufferedReader(new FileReader("c:\\test.txt"));
        // Start a line count and declare a string to hold our current line.
        int linecount = 0;
        String line;

        System.out.println("Searching for " + searchString);

        while (( line = bf.readLine()) != null)
        {
//            // Increment the count and find the index of the word
//            linecount++;
//            int indexfound = line.indexOf(args[0]);
//            // If greater than -1, means we found the word
//            if (indexfound > -1) {
//                System.out.println("Word was found at position " + indexfound + " on line " + linecount);
//            }
        }
        // Close the file after done searching
        bf.close();


    }
}



