package org.example;

import com.itextpdf.text.*;
import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

public class Poi {
    public static void main(String[] args) {

        try (FileInputStream docxInput = new FileInputStream("./files/test2.docx");
             FileOutputStream pdfOutput = new FileOutputStream("./output/output.pdf")) {

            XWPFDocument docx = new XWPFDocument(docxInput);
            Document pdf = getDocument(docx);

            Double doubleDefaultFontSize = docx.getStyles().getDefaultRunStyle().getFontSizeAsDouble();
            int defaultFontSize = doubleDefaultFontSize != null ? doubleDefaultFontSize.intValue() : 11;
            PdfWriter writer = PdfWriter.getInstance(pdf, pdfOutput);
            pdf.open();
            // Загрузка шрифта Arial
            BaseFont arialBaseFont = BaseFont.createFont("./ARIAL.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font arialFont = new Font(arialBaseFont, defaultFontSize, Font.NORMAL);

            for (IBodyElement element : docx.getBodyElements()) {
                if (element instanceof XWPFParagraph) {
                    XWPFParagraph para = (XWPFParagraph) element;
                    processParagraph(para, pdf, arialFont, defaultFontSize);
                } else if (element instanceof XWPFTable) {
                    XWPFTable table = (XWPFTable) element;
                    processTable(table, pdf, arialFont, defaultFontSize);
                }
            }

            pdf.close();
            docx.close();

            System.out.println("DOCX with images and tables converted to PDF successfully.");
        } catch (IOException | DocumentException e) {
            e.printStackTrace();
        }
    }

    private static void processParagraph(XWPFParagraph para, Document pdf, Font arialFont, int defaultFontSize) throws DocumentException, IOException {
        StringBuilder sb = new StringBuilder();
        Font font = arialFont;
        if (para.getRuns().isEmpty()) {
            pdf.add(new Paragraph(" ", arialFont));
            return;
        }


        for (XWPFRun run : para.getRuns()) {
            String text = run.getText(0);
            if (text != null) {
                float fontSize = run.getFontSizeAsDouble() != null ? run.getFontSizeAsDouble().floatValue() : 11f;
                if (fontSize == -1.0) fontSize = defaultFontSize;

                int style = Font.NORMAL;
                if (run.isBold()) style |= Font.BOLD;
                if (run.isItalic()) style |= Font.ITALIC;

                font = new Font(arialFont.getBaseFont(), fontSize, style);

                sb.append(text);
            }
            if (run.getEmbeddedPictures().size() > 0) {
                for (XWPFPicture picture : run.getEmbeddedPictures()) {
                    XWPFPictureData pictureData = picture.getPictureData();
                    byte[] pictureBytes = pictureData.getData();
                    Image img = Image.getInstance(pictureBytes);

                    // Получение размеров изображения из XWPFPicture
                    int width = (int) (picture.getCTPicture().getSpPr().getXfrm().getExt().getCx() / 9525);
                    int height = (int) (picture.getCTPicture().getSpPr().getXfrm().getExt().getCy() / 9525);

                    // Установка абсолютной позиции изображения, если необходимо
                    if (picture.getCTPicture().getSpPr().getXfrm().getOff() != null) {
                        long posX = (long) picture.getCTPicture().getSpPr().getXfrm().getOff().getX() / 9525;
                        long posY = (long) picture.getCTPicture().getSpPr().getXfrm().getOff().getY() / 9525;
                        img.setAbsolutePosition(posX, PageSize.A4.getHeight() - posY - height);
                    } else {
                        img.scaleToFit(width, height);
                    }
                    pdf.add(img);
                }
            }
        }
        Paragraph p = new Paragraph(sb.toString(), font);
        int spacingAfter = para.getSpacingAfter();
        int spacingBefore = para.getSpacingBefore();
        double spacingBetween = para.getSpacingBetween();
        para.setSpacingBetween(spacingBetween);

        if (spacingAfter == -1) {
            if (para.getSpacingLineRule().equals(LineSpacingRule.AUTO)) {
                p.setSpacingAfter(1f);
            }
        } else {
            p.setSpacingAfter(spacingAfter);
        }
        if (spacingBefore == -1) {
            if (para.getSpacingLineRule().equals(LineSpacingRule.AUTO)) {
                p.setSpacingBefore(1f);
            }
        } else {
            p.setSpacingBefore(spacingBefore);
        }

        pdf.add(p);
    }

    private static void processTable(XWPFTable table, Document pdf, Font arialFont, int defaultFontSize) throws DocumentException {
        int numCols = table.getRow(0).getTableCells().size();
        table.getLeftBorderSpace();
        PdfPTable pdfTable = new PdfPTable(numCols);
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                PdfPCell pdfCell = new PdfPCell();
                for (XWPFParagraph para : cell.getParagraphs()) {
                    Paragraph p = new Paragraph();
                    for (XWPFRun run : para.getRuns()) {
                        String text = run.getText(0);
                        if (text != null) {
                            int fontSize = run.getFontSize();
                            if (fontSize == -1) fontSize = defaultFontSize;

                            int style = Font.NORMAL;
                            if (run.isBold()) style |= Font.BOLD;
                            if (run.isItalic()) style |= Font.ITALIC;

                            Font font = new Font(arialFont.getBaseFont(), fontSize, style);
                            p.add(new Phrase(text, font));
                        }
                    }
                    pdfCell.addElement(p);
                }
                pdfTable.addCell(pdfCell);
            }
        }

        pdf.add(pdfTable);
    }

    private static Document getDocument(XWPFDocument docx) {
        CTSectPr sectPr = docx.getDocument().getBody().getSectPr();
        CTPageMar pageMar = sectPr.getPgMar();
        float leftMargin = ((BigInteger) pageMar.getLeft()).intValue() / 20f;
        float rightMargin = ((BigInteger) pageMar.getRight()).intValue() / 20f;
        float topMargin = ((BigInteger) pageMar.getTop()).intValue() / 20f;
        float bottomMargin = ((BigInteger) pageMar.getBottom()).intValue() / 20f;
        return new Document(PageSize.A4, leftMargin, rightMargin, topMargin, bottomMargin);
    }
}
