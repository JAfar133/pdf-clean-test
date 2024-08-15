package org.example;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.docx4j.Docx4J;
import org.docx4j.XmlUtils;
import org.docx4j.convert.out.ConversionFeatures;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.dml.wordprocessingDrawing.Anchor;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.P;
import org.docx4j.wml.R;

import java.io.*;
import java.util.*;

public class Docx4jTest {
    public static void main(String[] args) throws Exception {
        // Загрузите DOCX файл
        File inputFile = new File("./files/otchet.docx");
        InputStream is = new FileInputStream(inputFile);
        XWPFDocument document = new XWPFDocument(is);
        removeHeadersAndFooters2(document);
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        document.write(byteArrayOutputStream);
        byteArrayOutputStream.close();

        ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());
        XWPFDocument document2 = new XWPFDocument(byteArrayInputStream);
        try (FileOutputStream out = new FileOutputStream("./files/otchet_no_headers.docx")) {
            document2.write(out);
        }
    }

    private static void removeHeadersAndFooters2(XWPFDocument document) {
        // Удаление всех заголовков
        for (XWPFHeader header : document.getHeaderList()) {
            header.clearHeaderFooter();
        }
        // Удаление всех нижних колонтитулов
        for (XWPFFooter footer : document.getFooterList()) {
            footer.clearHeaderFooter();
        }
    }

    private static Mapper loadFontMapper() {
        Mapper fontMapper = new IdentityPlusMapper();
        fontMapper.put("Lishu", PhysicalFonts.get("LiSu"));
        fontMapper.put("Song", PhysicalFonts.get("SimSun"));
        fontMapper.put("Microsoft Ya", PhysicalFonts.get("Microsoft Yahei"));
        fontMapper.put("Black body", PhysicalFonts.get("SimHei"));
        fontMapper.put("", PhysicalFonts.get("KaiTi"));
        fontMapper.put("New Song", PhysicalFonts.get("NSimSun"));
        fontMapper.put("Huawen", PhysicalFonts.get("STXingkai"));
        fontMapper.put("Huawen imitation Song", PhysicalFonts.get("STFangsong"));
        fontMapper.put("Imitation Song", PhysicalFonts.get("FangSong"));
        fontMapper.put("Young circle", PhysicalFonts.get("YouYuan"));
        fontMapper.put("Huawen Song", PhysicalFonts.get("STSong"));
        fontMapper.put("Chinese Song Song", PhysicalFonts.get("STZhongsong"));
        fontMapper.put("Waiting", PhysicalFonts.get("SimSun"));
        fontMapper.put("Waiting Light", PhysicalFonts.get("SimSun"));
        fontMapper.put("Huawen Amber", PhysicalFonts.get("STHupo"));
        fontMapper.put("Huawen Lishu", PhysicalFonts.get("STLiti"));
        fontMapper.put("Huawen Xin Wei", PhysicalFonts.get("STXinwei"));
        fontMapper.put("Huang Wencai Yun", PhysicalFonts.get("STCaiyun"));
        fontMapper.put("Fang Zheng Yao", PhysicalFonts.get("FZYaoti"));
        fontMapper.put("Fang Zhengshu", PhysicalFonts.get("FZShuTi"));
        fontMapper.put("Huawen fine black", PhysicalFonts.get("STXihei"));
        fontMapper.put("Song expansion", PhysicalFonts.get("simsun-extB"));
        fontMapper.put("Imitation Song _GB2312", PhysicalFonts.get("FangSong_GB2312"));
        fontMapper.put("New Mind", PhysicalFonts.get("SimSun"));
        fontMapper.put("Arial", PhysicalFonts.get("Arial"));

        return fontMapper;
    }

    private static void processOverlapImages(WordprocessingMLPackage wordMLPackage) {
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        org.docx4j.wml.Document wmlDocumentEl = documentPart.getJaxbElement();
        Body body =  wmlDocumentEl.getBody();
        List<Object> bodyChildren = body.getContent();
        for (Object obj : bodyChildren) {
            Object unwrapped = XmlUtils.unwrap(obj);

            if (unwrapped instanceof P) {
                P paragraph = (P) unwrapped;
                List<Object> paragraphChildren = paragraph.getContent();

                for (Object paragraphObj : paragraphChildren) {
                    Object paragraphUnwrapped = XmlUtils.unwrap(paragraphObj);

                    if (paragraphUnwrapped instanceof R) {
                        R run = (R) paragraphUnwrapped;
                        List<Object> runChildren = run.getContent();

                        Iterator<Object> runIterator = runChildren.iterator();
                        while (runIterator.hasNext()) {
                            Object runObj = runIterator.next();
                            Object runUnwrapped = XmlUtils.unwrap(runObj);

                            if (runUnwrapped instanceof Drawing) {
                                Drawing drawing = (Drawing) runUnwrapped;

                                List<Object> anchorOrInline = drawing.getAnchorOrInline();
                                for (Object anchorOrInlineObj : anchorOrInline) {
                                    if (anchorOrInlineObj instanceof Anchor) {
                                        Anchor anchor = (Anchor) anchorOrInlineObj;
                                        if (anchor.isAllowOverlap()) {
                                            runIterator.remove();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    private static void processEmptyTableRows(WordprocessingMLPackage wordMLPackage) {
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        org.docx4j.wml.Document wmlDocumentEl = documentPart.getJaxbElement();
        Body body = wmlDocumentEl.getBody();
        List<Object> bodyChildren = body.getContent();
        for (Object obj : bodyChildren) {
            Object unwrapped = XmlUtils.unwrap(obj);

            if (unwrapped instanceof org.docx4j.wml.Tbl) {
                org.docx4j.wml.Tbl table = (org.docx4j.wml.Tbl) unwrapped;
                List<Object> tableChildren = table.getContent();
                List<Object> rowsToRemove = new ArrayList<>();
                for (Object tableObj : tableChildren) {
                    Object tableUnwrapped = XmlUtils.unwrap(tableObj);

                    if (tableUnwrapped instanceof org.docx4j.wml.Tr) {
                        org.docx4j.wml.Tr tableRow = (org.docx4j.wml.Tr) tableUnwrapped;
                        List<Object> rowChildren = tableRow.getContent();

                        boolean hasTableCell = false;
                        for (Object rowObj : rowChildren) {
                            Object rowUnwrapped = XmlUtils.unwrap(rowObj);

                            if (rowUnwrapped instanceof org.docx4j.wml.Tc) {
                                hasTableCell = true;
                                break;
                            }
                        }

                        if (!hasTableCell) {
                            rowsToRemove.add(tableRow);
                        }
                    }
                }
                if (!rowsToRemove.isEmpty()) {
                    tableChildren.removeAll(rowsToRemove);
                }
            }
        }
    }


}
