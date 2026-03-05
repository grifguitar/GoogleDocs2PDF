package me.kgrigoriy;

import com.microsoft.playwright.*;
import com.microsoft.playwright.options.LoadState;
import com.microsoft.playwright.options.ScreenshotType;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import net.sourceforge.tess4j.ITessAPI;
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.Word;

import org.apache.commons.io.FileUtils;
import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.pdmodel.graphics.state.RenderingMode;

import java.awt.Rectangle;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.regex.*;

import javax.imageio.ImageIO;

public class Main {

    public static void main(String[] args) {
        App.run(args);
    }

    public enum Google {
        Presentation(1920, 1080,
                "https://docs.google.com/presentation/d/",
                "docs\\.google\\.com/presentation/d/([a-zA-Z0-9_-]+)"),
        Document(1000, 1150,
                "https://docs.google.com/document/d/",
                "docs\\.google\\.com/document/d/([a-zA-Z0-9_-]+)");

        static final int DOC_WIDTH = 820;
        static final int DOC_SCROLL_STEP = 1133;
        static final String PREVIEW = "/preview";
        static final double SCALE = 3.0;
        static final int MAX_PAGES = 1000;

        final int width, height;
        final String prefix, regex;

        Google(int w, int h, String p, String r) {
            width = w;
            height = h;
            prefix = p;
            regex = r;
        }
    }

    public static record Link(Google t, String url) {
    }

    public static class App extends Application {

        public static void run(String[] args) {
            launch(args);
        }

        private TextField urlField;
        private TextField outputField;
        private Button convertButton;
        private CheckBox ocrCheckBox;
        private CheckBox onlyTextCheckBox;
        private TextArea logArea;
        private Stage primaryStage;

        private Tesseract tesseract;
        private Path tmpDir;

        @Override
        public void start(Stage stage) {
            this.primaryStage = stage;
            stage.setTitle("GoogleDocs2PDF by KGrigoriy for Anna💕 v1.0");

            InputStream iconStream = getClass().getResourceAsStream("/icon.png");
            if (iconStream != null) {
                stage.getIcons().add(new Image(iconStream));
            }

            Label urlLabel = new Label("Ссылка:");
            urlLabel.setMinWidth(80);
            urlField = new TextField();
            urlField.setPromptText("https://docs.google.com/");
            HBox.setHgrow(urlField, Priority.ALWAYS);
            HBox urlRow = new HBox(8, urlLabel, urlField);
            urlRow.setAlignment(Pos.CENTER_LEFT);

            Label outLabel = new Label("Сохранить:");
            outLabel.setMinWidth(80);
            outputField = new TextField();
            outputField.setPromptText("output.pdf");
            HBox.setHgrow(outputField, Priority.ALWAYS);
            Button browseBtn = new Button("Обзор");
            browseBtn.setOnAction(e -> {

                FileChooser fc = new FileChooser();
                fc.setTitle("Сохранить PDF как");
                fc.getExtensionFilters().add(
                        new FileChooser.ExtensionFilter("PDF файлы", "*.pdf"));
                fc.setInitialFileName("output.pdf");
                File file = fc.showSaveDialog(primaryStage);
                if (file != null)
                    outputField.setText(file.getAbsolutePath());

            });
            HBox outRow = new HBox(8, outLabel, outputField, browseBtn);
            outRow.setAlignment(Pos.CENTER_LEFT);

            ocrCheckBox = new CheckBox("OCR");
            ocrCheckBox.setSelected(true);

            onlyTextCheckBox = new CheckBox("Text only");
            onlyTextCheckBox.setSelected(false);
            onlyTextCheckBox.disableProperty().bind(ocrCheckBox.selectedProperty().not());

            convertButton = new Button("Конвертировать");
            convertButton.setDefaultButton(true);
            convertButton.setOnAction(e -> {

                String rawUrl = urlField.getText().trim();
                String outputRaw = outputField.getText().trim();

                if (rawUrl.isEmpty()) {
                    log("Введите ссылку");
                    return;
                }
                if (outputRaw.isEmpty()) {
                    log("Введите имя файла для сохранения");
                    return;
                }

                Link link = parseLink(rawUrl);
                if (link == null) {
                    log("Некорректная ссылка: поддерживаются Google Slides и Google Docs");
                    return;
                }

                final String outputPath = outputRaw.toLowerCase().endsWith(".pdf")
                        ? outputRaw
                        : outputRaw + ".pdf";

                if (!outputRaw.equals(outputPath))
                    Platform.runLater(() -> outputField.setText(outputPath));

                final boolean useOcr = ocrCheckBox.isSelected();
                final boolean onlyText = onlyTextCheckBox.isSelected() && useOcr;

                convertButton.setDisable(true);
                logArea.clear();

                Thread thread = new Thread(() -> {
                    try {
                        log("Тип документа: " + link.t().name());
                        ensureChromiumInstalled();
                        if (useOcr)
                            ensureTesseractReady();
                        try (PDDocument pdf = new PDDocument();
                                Playwright playwright = Playwright.create();
                                Browser browser = playwright.chromium().launch(
                                        new BrowserType.LaunchOptions().setHeadless(true));
                                BrowserContext ctx = browser.newContext(
                                        new Browser.NewContextOptions().setDeviceScaleFactor(Google.SCALE));
                                Page page = ctx.newPage()) {
                            parsePage(page, link, pdf, useOcr, onlyText);
                            pdf.save(outputPath);
                            log("PDF сохранён: " + outputPath);
                        }
                    } catch (InterruptedException ex) {
                        Thread.currentThread().interrupt();
                        log("Конвертация прервана");
                    } catch (Exception ex) {
                        log("Ошибка: " + ex);
                    } finally {
                        Platform.runLater(() -> convertButton.setDisable(false));
                    }
                }, "converter-thread");
                thread.setDaemon(true);
                thread.start();

            });
            HBox btnRow = new HBox(12, ocrCheckBox, onlyTextCheckBox, convertButton);
            btnRow.setAlignment(Pos.CENTER_RIGHT);

            logArea = new TextArea();
            logArea.setEditable(false);
            logArea.setPrefHeight(230);
            logArea.setWrapText(true);
            logArea.setStyle("-fx-font-family: 'Courier New', monospace; -fx-font-size: 12px;");
            VBox.setVgrow(logArea, Priority.ALWAYS);

            VBox root = new VBox(10, urlRow, outRow, btnRow, new Label("Лог:"), logArea);
            root.setPadding(new Insets(16));

            stage.setScene(new Scene(root, 680, 440));
            stage.setMinWidth(500);
            stage.setMinHeight(380);
            stage.show();
        }

        @Override
        public void stop() {
            if (tmpDir != null)
                FileUtils.deleteQuietly(tmpDir.toFile());
        }

        private void log(String msg) {
            Platform.runLater(() -> logArea.appendText(msg + "\n"));
        }

        private Link parseLink(String link) {
            Matcher slides = Pattern.compile(Google.Presentation.regex).matcher(link);
            if (slides.find())
                return new Link(Google.Presentation,
                        Google.Presentation.prefix + slides.group(1) + Google.PREVIEW);
            Matcher docs = Pattern.compile(Google.Document.regex).matcher(link);
            if (docs.find())
                return new Link(Google.Document,
                        Google.Document.prefix + docs.group(1) + Google.PREVIEW);
            return null;
        }

        private Path getCachePath() {
            String os = System.getProperty("os.name").toLowerCase();
            String home = System.getProperty("user.home");
            if (os.contains("win"))
                return Paths.get(home, "AppData", "Local", "ms-playwright");
            if (os.contains("mac") || os.contains("darwin"))
                return Paths.get(home, "Library", "Caches", "ms-playwright");
            return Paths.get(home, ".cache", "ms-playwright");
        }

        private void ensureChromiumInstalled() throws IOException, InterruptedException {
            if (Files.exists(getCachePath())) {
                log("Chromium уже установлен");
                return;
            }
            log("Устанавливаю Chromium");
            ProcessBuilder pb = new ProcessBuilder(
                    "java", "-cp", System.getProperty("java.class.path"),
                    "com.microsoft.playwright.CLI", "install", "chromium");
            pb.inheritIO();
            int code = pb.start().waitFor();
            if (code != 0)
                throw new IOException("Код выхода: " + code);
            log("Chromium установлен");
        }

        private void ensureTesseractReady() throws IOException {
            if (tesseract != null) {
                return;
            }
            log("Подготовка OCR");
            if (tmpDir != null) {
                FileUtils.deleteQuietly(tmpDir.toFile());
                tmpDir = null;
            }
            tmpDir = Files.createTempDirectory("gdocs2pdf");
            Path data = tmpDir.resolve("tessdata");
            Files.createDirectory(data);
            for (String f : List.of("rus.traineddata", "eng.traineddata", "osd.traineddata")) {
                InputStream is = getClass().getResourceAsStream("/tessdata/" + f);
                if (is == null)
                    throw new IOException("Отсутствует: /tessdata/" + f);
                Files.copy(is, data.resolve(f));
            }
            log("OCR загружен");
            tesseract = new Tesseract();
            tesseract.setDatapath(data.toAbsolutePath().toString());
            tesseract.setLanguage("rus+eng");
            tesseract.setPageSegMode(1);
            tesseract.setOcrEngineMode(1);
            tesseract.setVariable("user_defined_dpi", String.valueOf((int) (96 * Google.SCALE)));
            log("OCR готов");
        }

        private void parsePage(Page page, Link link, PDDocument pdf, boolean ocr, boolean txt) throws IOException {
            page.setViewportSize(link.t().width, link.t().height);

            PDFont font = null;
            if (ocr) {
                InputStream is = getClass().getResourceAsStream("/fonts/DejaVuSans.ttf");
                if (is == null)
                    throw new IOException("Отсутствует: /fonts/DejaVuSans.ttf");
                font = PDType0Font.load(pdf, is);
            }

            byte[] prev = null;
            int saved = 0;

            for (int num = 1; num <= Google.MAX_PAGES; num++) {
                try {
                    if (link.t() == Google.Presentation || num == 1) {
                        String anchor = (link.t() == Google.Presentation) ? "#slide=id.p" + num : "";
                        page.navigate(link.url() + anchor);
                        page.waitForLoadState(LoadState.NETWORKIDLE,
                                new Page.WaitForLoadStateOptions());
                    }

                    byte[] shot = page.screenshot(
                            new Page.ScreenshotOptions()
                                    .setType(ScreenshotType.PNG).setFullPage(false));

                    if (link.t() == Google.Document) {
                        BufferedImage original = ImageIO.read(new ByteArrayInputStream(shot));
                        if (original == null)
                            throw new IOException("Не удалось декодировать скриншот");
                        int cropWidth = Math.min((int) (Google.DOC_WIDTH * Google.SCALE), original.getWidth());
                        int x = Math.max(0, (original.getWidth() - cropWidth) / 2);
                        BufferedImage cropped = original.getSubimage(x, 0, cropWidth, original.getHeight());
                        ByteArrayOutputStream baos = new ByteArrayOutputStream();
                        if (!ImageIO.write(cropped, "png", baos))
                            throw new IOException("PNG writer недоступен");
                        shot = baos.toByteArray();
                    }

                    if (Arrays.equals(prev, shot)) {
                        log("Страниц всего: " + saved);
                        return;
                    }

                    try {
                        addPage(pdf, shot, num, font, txt);
                        saved++;
                        log("Страница: " + num);
                    } catch (Exception e) {
                        log("Страница " + num + " пропущена: " + e);
                    }

                    prev = shot;

                    if (link.t() == Google.Document) {
                        page.mouse().wheel(0, Google.DOC_SCROLL_STEP);
                        page.waitForLoadState(LoadState.NETWORKIDLE,
                                new Page.WaitForLoadStateOptions());
                    }

                } catch (Exception e) {
                    if (num == 1)
                        throw new IOException("Ошибка загрузки первой страницы", e);
                    log("Страница " + num + " недоступна, пропускаю: " + e);
                    if (link.t() == Google.Presentation && prev != null) {
                        log("Страниц всего: " + saved);
                        return;
                    }
                }
            }
            log("ВНИМАНИЕ: достигнут лимит в " + Google.MAX_PAGES + " страниц");
        }

        private void addPage(PDDocument pdf, byte[] shot, int num, PDFont font, boolean txt) throws IOException {
            BufferedImage image = ImageIO.read(new ByteArrayInputStream(shot));
            if (image == null)
                throw new IOException("Ошибка декодирования скриншота на странице: " + num);
            PDImageXObject pdImg = PDImageXObject.createFromByteArray(pdf, shot, "p" + num);
            float pageWidth = pdImg.getWidth() / (float) Google.SCALE;
            float pageHeight = pdImg.getHeight() / (float) Google.SCALE;
            PDPage pdfPage = new PDPage(new PDRectangle(pageWidth, pageHeight));
            pdf.addPage(pdfPage);

            try (PDPageContentStream stream = new PDPageContentStream(pdf, pdfPage)) {
                if (!txt)
                    stream.drawImage(pdImg, 0, 0, pageWidth, pageHeight);

                if (font == null || tesseract == null)
                    return;

                List<Word> words;
                try {
                    words = tesseract.getWords(image, ITessAPI.TessPageIteratorLevel.RIL_WORD);
                } catch (Exception e) {
                    log("OCR ошибка на странице " + num + ": " + e);
                    return;
                }
                if (words == null || words.isEmpty())
                    return;

                stream.setRenderingMode(txt ? RenderingMode.FILL : RenderingMode.NEITHER);

                float scaleX = pageWidth / image.getWidth();
                float scaleY = pageHeight / image.getHeight();
                float lastFontSize = -1f;
                float lastHScale = -1f;

                List<List<Word>> lines = groupIntoLines(words);

                List<Float> fontSizes = new ArrayList<>();
                for (List<Word> line : lines) {
                    float sz = 0;
                    for (Word w : line) {
                        float h = w.getBoundingBox().height * scaleY;
                        sz = Math.max(sz, Math.max(h * 0.85f, 1f));
                    }
                    fontSizes.add(sz);
                }
                Collections.sort(fontSizes);
                float pageFontSize = fontSizes.get(fontSizes.size() / 2);

                for (List<Word> line : lines) {
                    float lineBottom = 0;
                    for (Word w : line)
                        lineBottom = Math.max(lineBottom,
                                (w.getBoundingBox().y + w.getBoundingBox().height) * scaleY);
                    float pdfLineY = pageHeight - lineBottom;

                    for (Word word : line) {
                        String text = word.getText().trim();
                        if (text.isEmpty())
                            continue;

                        Rectangle bbox = word.getBoundingBox();
                        float pdfX = bbox.x * scaleX;
                        float pdfWidth = bbox.width * scaleX;

                        try {
                            float textWidth = font.getStringWidth(text) / 1000f * pageFontSize;
                            if (textWidth <= 0)
                                continue;
                            float hScale = Math.max(10f, Math.min((pdfWidth / textWidth) * 100f, 300f));

                            if (pageFontSize != lastFontSize || hScale != lastHScale) {
                                stream.setFont(font, pageFontSize);
                                stream.setHorizontalScaling(hScale);
                                lastFontSize = pageFontSize;
                                lastHScale = hScale;
                            }

                            stream.beginText();
                            stream.newLineAtOffset(pdfX, pdfLineY);
                            try {
                                stream.showText(text);
                            } catch (Exception e) {
                                log("OCR ошибка при обработке " + text + " на странице " + num + ": " + e);
                            } finally {
                                stream.endText();
                            }
                        } catch (Exception ignored) {
                        }
                    }
                }
            }
        }

        private List<List<Word>> groupIntoLines(List<Word> words) {
            List<List<Word>> lines = new ArrayList<>();
            List<int[]> spans = new ArrayList<>();

            for (Word word : words) {
                Rectangle bbox = word.getBoundingBox();
                boolean added = false;
                for (int i = 0; i < lines.size(); i++) {
                    int[] span = spans.get(i);
                    if (bbox.y < span[1] && bbox.y + bbox.height > span[0]) {
                        lines.get(i).add(word);
                        span[0] = Math.min(span[0], bbox.y);
                        span[1] = Math.max(span[1], bbox.y + bbox.height);
                        added = true;
                        break;
                    }
                }
                if (!added) {
                    List<Word> newLine = new ArrayList<>();
                    newLine.add(word);
                    lines.add(newLine);
                    spans.add(new int[] { bbox.y, bbox.y + bbox.height });
                }
            }
            return lines;
        }

    }

}
