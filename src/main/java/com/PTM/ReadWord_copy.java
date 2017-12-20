package com.PTM;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

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
        Iterator<IBodyElement> bodyElementIterator = doc.getBodyElementsIterator();
        while (bodyElementIterator.hasNext()) {

            IBodyElement element = bodyElementIterator.next();

            if ("TABLE".equalsIgnoreCase(element.getElementType().name())) {
                List<XWPFTable> tables = element.getBody().getTables();

                for (XWPFTable table : tables) {

                    System.out.println("In Table and First Cell is: =================" + table.getRow(0).getCell(0).getText());
//                    System.out.println(table.getText());
                    if (table.getRow(0).getCell(0).getText().equals("Tag")){
                        System.out.println("Table is considered");
                        System.out.println(table.getText());
//                        continue;
                    }
//                    System.out.println("==============================================================================================================");

//                    for (XWPFTableRow row : table.getRows()) {
//
//
//                        if (row.getCell(0).getText().isEmpty()) {
//                            System.out.println("Skip cell and continue in table");
//                            continue;
//                        }
//                        if (!row.getCell(0).getText().startsWith("0")) {
//                            System.out.println("Skip cell and continue in table");
//                            continue;
//                        }
//
//                        System.out.println("+++++++"+row.getCell(0).getText());
//                        //cellOne =
//                    }
                }
            }
        }
    }
}



