package org.example;

import org.apache.commons.lang3.StringUtils;
import org.apache.lucene.search.spell.LevenshteinDistance;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.pdfbox.text.TextPosition;

import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.IOException;
import java.util.*;

public class PdfBOXX {
    public static void main(String[] args) throws IOException {
        // Открываем существующий PDF-документ
        PDDocument document = PDDocument.load(new File("./files/Maples+Group+-+Transfer+Agent+FAQ+-+August+2023.pdf"));

        for (int pageIndex = 0; pageIndex < document.getNumberOfPages(); pageIndex++) {
            PDPage page = document.getPage(pageIndex);
            PDRectangle mediaBox = page.getMediaBox();
            float lowerLeftY = mediaBox.getLowerLeftY();
            float upperRightY = mediaBox.getUpperRightY();
            float lowerLeftX = mediaBox.getLowerLeftX();
            float upperRightX = mediaBox.getUpperRightX();

            // Определяем области колонтитулов в виде Rectangle2D
            Rectangle2D footerRegion = new Rectangle2D.Float(lowerLeftX, upperRightY - 1000, upperRightX - lowerLeftX, 1000);
            Rectangle2D headerRegion = new Rectangle2D.Float(lowerLeftX, lowerLeftY, upperRightX - lowerLeftX, 1000);

            Set<String> clusters = findRepetitiveLinesAndPatterns(document, headerRegion);
            clusters.addAll(findRepetitiveLinesAndPatterns(document, footerRegion));

            // Извлекаем текст без колонтитулов
            String textWithoutHeadersFooters = removeRepetitiveGroups(document, pageIndex + 1, clusters, headerRegion, footerRegion);
            System.out.println(textWithoutHeadersFooters);
        }

        // Сохраняем изменённый документ
        document.save("./output/output-test.pdf");
        document.close();
    }

    private static String removeRepetitiveGroups(PDDocument document, int pageIndex, Set<String> clusters, Rectangle2D headerRegion, Rectangle2D footerRegion) throws IOException {
        PDFTextStripper stripper = new PDFTextStripper() {
            @Override
            protected void writeString(String string, List<TextPosition> textPositions) throws IOException {
                boolean isWrite = true;
                for (TextPosition text : textPositions) {
                    float x = text.getX();
                    float y = text.getY();
                    if (clusters.contains(string.trim()) && (headerRegion.contains(x, y) || footerRegion.contains(x, y))) {
                        isWrite = false;
                        break;
                    }
                }
                if (isWrite) {
                    super.writeString(string, textPositions);
                }
            }
        };

        stripper.setStartPage(pageIndex);
        stripper.setEndPage(pageIndex);
        return stripper.getText(document).trim();
    }

    private static Set<String> findRepetitiveLinesAndPatterns(PDDocument document, Rectangle2D region) throws IOException {
        Map<String, Integer> commonLines = new HashMap<>();
        PDFTextStripperByArea stripperByArea = new PDFTextStripperByArea();
        stripperByArea.addRegion("region", region);

        for (int i = 1; i <= document.getNumberOfPages(); i++) {
            stripperByArea.extractRegions(document.getPage(i - 1));
            String text = stripperByArea.getTextForRegion("region");
            String[] lines = text.split("\r?\n");

            for (String line : lines) {
                String cleanedLine = line.trim();
                if (!cleanedLine.isEmpty()) {
                    commonLines.put(cleanedLine, commonLines.getOrDefault(cleanedLine, 0) + 1);
                }
            }
        }

        return findDuplicateSets(commonLines, document.getNumberOfPages());
    }
    private static Set<String> findDuplicateSets(Map<String, Integer> lineCounts, int totalPages) {
        final float SIMILARITY_THRESHOLD = 0.8f;
        final float FREQUENCY_THRESHOLD = 0.7f;

        LevenshteinDistance levenshteinDistance = new LevenshteinDistance();
        Set<String> linesToRemove = new HashSet<>();
        List<String> lines = new ArrayList<>(lineCounts.keySet());
        boolean[] visited = new boolean[lines.size()];

        for (int i = 0; i < lines.size(); i++) {
            if (visited[i]) continue;
            String line = lines.get(i);

            if ((float) (lineCounts.get(line) / totalPages) >= FREQUENCY_THRESHOLD || StringUtils.isNumeric(line)) {
                visited[i] = true;
                linesToRemove.add(line);
                continue;
            }

            List<String> patterns = new ArrayList<>();
            patterns.add(line);
            int patternCount = lineCounts.get(line);
            for (int j = i + 1; j < lines.size(); j++) {
                if (visited[j]) continue;
                String otherPattern = lines.get(j);
                if (levenshteinDistance.getDistance(line, otherPattern) >= SIMILARITY_THRESHOLD) {
                    patterns.add(otherPattern);
                    patternCount += lineCounts.get(otherPattern);
                    visited[j] = true;
                }
            }
            if ((float) (patternCount / totalPages) > FREQUENCY_THRESHOLD) {
                linesToRemove.addAll(patterns);
            }

        }

        return linesToRemove;
    }


    private static String removeEmptyLines(String text) {
        String[] lines = text.split("\r?\n");
        StringBuilder trimmedText = new StringBuilder();

        for (String line : lines) {
            if (!line.trim().isEmpty()) {
                trimmedText.append(line).append("\n");
            }
        }

        return trimmedText.toString().trim();
    }
}
