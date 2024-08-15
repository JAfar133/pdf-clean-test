package org.example;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.math3.util.Pair;
import org.apache.lucene.search.spell.LevenshteinDistance;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.pdfbox.text.TextPosition;

import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.IOException;
import java.util.*;

public class pdfBoxxTests {

    public static void main(String[] args) throws IOException {
        File dir = new File("./test");
        File[] filesArray = dir.listFiles((dir1, name) -> name.toLowerCase().endsWith(".pdf"));
        if (filesArray == null) {
            System.out.println("No PDF files found in the directory.");
            return;
        }

        List<String> files = new ArrayList<>();
        for (File file : filesArray) {
            files.add(file.getAbsolutePath());
        }

        Map<String, Integer> clusterFrequency = new HashMap<>();
        int totalPageCount = 0;
        for (String file : files) {
            PDDocument document = PDDocument.load(new File(file));
            totalPageCount += document.getNumberOfPages();
            Map<String, LineInfo> clusters = run(document);
            for (String cluster : clusters.keySet()) {
                clusterFrequency.put(cluster, clusters.get(cluster).count + clusterFrequency.getOrDefault(cluster, 0));
            }
            document.close();
        }
        System.out.println("Files count: " + files.size());
        System.out.println("Total page count: " + totalPageCount);
        // Найти и вывести наиболее частые кластеры
        clusterFrequency.entrySet().stream()
                .sorted((e1, e2) -> e2.getValue().compareTo(e1.getValue()))
                .limit(100) // количество наиболее частых кластеров для вывода
                .forEach(e -> System.out.println("Line: '" + e.getKey() + "': " + e.getValue()));

    }

    private static Map<String, LineInfo> run(PDDocument document) throws IOException {
        PDRectangle mediaBox = document.getPage(0).getMediaBox();
        float lowerLeftY = mediaBox.getLowerLeftY();
        float upperRightY = mediaBox.getUpperRightY();
        float lowerLeftX = mediaBox.getLowerLeftX();
        float upperRightX = mediaBox.getUpperRightX();

        RectangleRegion footerRegion = new RectangleRegion(lowerLeftX, lowerLeftY, upperRightX - lowerLeftX, 75, "footerRegion");
        RectangleRegion headerRegion = new RectangleRegion(lowerLeftX, upperRightY - 75, upperRightX - lowerLeftX, 75, "headerRegion");
        RectangleRegion bodyRegion = new RectangleRegion(lowerLeftX, lowerLeftY + 75, upperRightX - lowerLeftX, upperRightY - lowerLeftY - 150, "bodyRegion");

        Map<String, LineInfo> clusters = findRepetitiveLinesAndPatterns(document, footerRegion, headerRegion, bodyRegion);

//        for (int pageIndex = 0; pageIndex < document.getNumberOfPages(); pageIndex++) {
//            String textWithoutHeadersFooters = removeRepetitiveGroups(document, pageIndex + 1, clusters);
//            System.out.println(textWithoutHeadersFooters);
//        }

        document.save("./output/output-test.pdf");
        document.close();

        return clusters;
    }

    static class RectangleRegion extends Rectangle2D.Float {

        public RectangleRegion(float x, float y, float w, float h, String regionStr) {
            super(x, y, w, h);
            this.regionStr = regionStr;
        }

        private String regionStr;

        public String getRegionStr() {
            return regionStr;
        }

        public void setRegionStr(String regionStr) {
            this.regionStr = regionStr;
        }
    }

    private static String removeRepetitiveGroups(PDDocument document, int pageIndex, Map<String, Set<RectangleRegion>> clusters) throws IOException {
        PDFTextStripper stripper = new PDFTextStripper() {
            @Override
            protected void writeString(String string, List<TextPosition> textPositions) throws IOException {
                Set<RectangleRegion> regions = clusters.get(string.trim());
                boolean isWrite = true;
                if (regions != null) {
                    for (TextPosition text : textPositions) {
                        float x = text.getX();
                        float y = text.getY();
                        for (RectangleRegion region : regions) {
                            if (region.contains(x, y)) {
                                isWrite = false;
                                break;
                            }
                        }
                        if (!isWrite) {
                            break;
                        }
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

    static class LineInfo {
        public Set<RectangleRegion> regions;
        public int count;

        public LineInfo(Set<RectangleRegion> regions, int count) {
            this.regions = regions;
            this.count = count;
        }
    }

    private static final List<String> EXTRA_BODY_LINES = List.of(
            "(No file attached)", "Please upload supporting document:"
    );
    private static Map<String, LineInfo> findRepetitiveLinesAndPatterns(PDDocument document, RectangleRegion footerRegion, RectangleRegion headerRegion, RectangleRegion bodyRegion) throws IOException {
        Map<String, LineInfo> commonLines = new HashMap<>();
        PDFTextStripperByArea stripperByArea = new PDFTextStripperByArea();
        stripperByArea.addRegion(footerRegion.getRegionStr(), footerRegion);
        stripperByArea.addRegion(headerRegion.getRegionStr(), headerRegion);
        stripperByArea.addRegion(bodyRegion.getRegionStr(), bodyRegion);

        for (int i = 1; i <= document.getNumberOfPages(); i++) {
            stripperByArea.extractRegions(document.getPage(i - 1));
//            String footerText = stripperByArea.getTextForRegion(footerRegion.getRegionStr());
//            String headerText = stripperByArea.getTextForRegion(headerRegion.getRegionStr());
            String bodyText = stripperByArea.getTextForRegion(bodyRegion.getRegionStr());
//            String[] footerLines = footerText.split("\r?\n");
//            String[] headerLines = headerText.split("\r?\n");
            String[] bodyLines = bodyText.split("\r?\n");

            List<Pair<String, RectangleRegion>> allLines = new ArrayList<>();
//            for (String line : footerLines) {
//                allLines.add(new Pair<>(line.trim(), footerRegion));
//            }
//            for (String line : headerLines) {
//                allLines.add(new Pair<>(line.trim(), headerRegion));
//            }
            for (String line : bodyLines) {
                allLines.add(new Pair<>(line, bodyRegion));
            }

            for (Pair<String, RectangleRegion> line : allLines) {
                String cleanedLine = line.getFirst().trim();
                if (!cleanedLine.isEmpty()) {
                    LineInfo lineInfo = commonLines.get(cleanedLine);
                    if (lineInfo == null) {
                        Set<RectangleRegion> regions = new HashSet<>();
                        regions.add(line.getSecond());
                        commonLines.put(cleanedLine, new LineInfo(regions, 1));
                    } else {
                        lineInfo.regions.add(line.getSecond());
                        lineInfo.count++;
                    }
                }
            }
        }

        return commonLines;
    }

    private static Map<String, Set<RectangleRegion>> findDuplicateSets(Map<String, LineInfo> lineCounts, int totalPages) {
        final float SIMILARITY_THRESHOLD = 0.0f;
        final float FREQUENCY_THRESHOLD = 0.0f;

        LevenshteinDistance levenshteinDistance = new LevenshteinDistance();
        Map<String, Set<RectangleRegion>> linesToRemove = new HashMap<>();
        List<String> lines = new ArrayList<>(lineCounts.keySet());
        boolean[] visited = new boolean[lines.size()];

        for (int i = 0; i < lines.size(); i++) {
            String line = lines.get(i);
            LineInfo lineInfo = lineCounts.get(line);
            if ((float) lineInfo.count / totalPages >= FREQUENCY_THRESHOLD
                    || (StringUtils.isNumeric(line) && !lineInfo.regions.stream().allMatch(r -> r.getRegionStr().equals("bodyRegion")))) {
                visited[i] = true;
                linesToRemove.put(line, lineInfo.regions);
                continue;
            }

            Map<String, Set<RectangleRegion>> patterns = new HashMap<>();
            patterns.put(line, lineInfo.regions);
            int patternCount = lineCounts.get(line).count;
            for (int j = i + 1; j < lines.size(); j++) {
                if (visited[j]) continue;
                String otherPattern = lines.get(j);
                LineInfo lineInfo1 = lineCounts.get(otherPattern);
                if (levenshteinDistance.getDistance(line, otherPattern) >= SIMILARITY_THRESHOLD) {
                    patterns.put(otherPattern, lineInfo1.regions);
                    patternCount += lineInfo1.count;
                    visited[j] = true;
                }
            }
            if ((float) patternCount / totalPages > FREQUENCY_THRESHOLD) {
                linesToRemove.putAll(patterns);
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
