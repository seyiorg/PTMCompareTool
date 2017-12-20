package com.PTM;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

public class ReadWord_copy {

    private String cellOne;
    private String cellTwo;
    private String cellThree;
    private String cellFour;
    public String searchString = cellOne + cellTwo + cellThree + cellFour;

    //Read table cell
    @Test
    public void readDocxTable() throws IOException, InvalidFormatException {

        FileInputStream fis = new FileInputStream("TC602 Vol 22.docx");
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(fis));
        int count = 0;
        int internalcount = 0;
        List<XWPFTable> tables = doc.getTables();
        for (XWPFTable table : tables) {
//            System.out.println(table.getText());
            count++;
            if (table.getRow(0).getCell(0).getText().contains("Tag")){
//                System.out.println(table.getText());
                internalcount++;
            }
//
//            for (XWPFTableRow row : table.getRows()) {
//
//
//            }

        }
        System.out.println("Total no of tables: "  + count);
        System.out.println("Total no of tables: "  + internalcount);

    }
}



