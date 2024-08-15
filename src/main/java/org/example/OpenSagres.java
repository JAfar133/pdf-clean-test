package org.example;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.docx4j.openpackaging.exceptions.Docx4JException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class OpenSagres {
    public static void main(String[] args) throws Exception {
        File inputFile = new File("./files/test2.docx");
        File outputFile = new File("./output/firstPage.docx");

        extractFirstPage(inputFile, outputFile);

        System.out.println("First page extracted and saved to " + outputFile.getAbsolutePath());
    }

    private static void extractFirstPage(File inputFile, File outputFile) throws IOException, Docx4JException {
        try (FileInputStream fis = new FileInputStream(inputFile);
             XWPFDocument document = new XWPFDocument(fis);
             XWPFDocument newDocument = new XWPFDocument()) {

            boolean firstPageFound = false;

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    if (run.getCTR().xmlText().contains("<w:lastRenderedPageBreak/>")) {
                        firstPageFound = true;
                        break;
                    }
                }

                if (firstPageFound) break;

                XWPFParagraph newParagraph = newDocument.createParagraph();
                newParagraph.getCTP().set(paragraph.getCTP());
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                newDocument.write(fos);
            }
        }
    }
}
