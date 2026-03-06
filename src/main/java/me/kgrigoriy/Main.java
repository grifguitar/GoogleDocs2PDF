package me.kgrigoriy;

import com.microsoft.playwright.*;
import com.microsoft.playwright.options.LoadState;
import com.microsoft.playwright.options.ScreenshotType;

import net.sourceforge.tess4j.ITessAPI;
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.Word;

import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.pdmodel.graphics.state.RenderingMode;

import javax.imageio.ImageIO;
import java.awt.Rectangle;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;
import java.util.concurrent.*;
import java.util.logging.*;
import java.util.regex.*;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;
import org.telegram.telegrambots.bots.TelegramLongPollingBot;
import org.telegram.telegrambots.meta.TelegramBotsApi;
import org.telegram.telegrambots.meta.api.methods.send.SendDocument;
import org.telegram.telegrambots.meta.api.methods.send.SendMessage;
import org.telegram.telegrambots.meta.api.objects.InputFile;
import org.telegram.telegrambots.meta.api.objects.Update;
import org.telegram.telegrambots.meta.api.objects.replykeyboard.ReplyKeyboardMarkup;
import org.telegram.telegrambots.meta.api.objects.replykeyboard.ReplyKeyboardRemove;
import org.telegram.telegrambots.meta.api.objects.replykeyboard.buttons.KeyboardRow;
import org.telegram.telegrambots.updatesreceivers.DefaultBotSession;

public class Main {
    private static final Logger LOG = Logger.getLogger(Main.class.getName());

    public static void main(String[] args) throws Exception {
        setupLogging();

        String username = "GoogleDocs2PDFBot";
        String token;
        try (BufferedReader in = new BufferedReader(
                new FileReader("token", StandardCharsets.UTF_8))) {
            token = in.readLine();
            if (token == null || token.strip().isEmpty())
                throw new IllegalStateException("Файл token пуст");
            token = token.strip();
        } catch (FileNotFoundException e) {
            throw new IllegalStateException("Файл 'token' не найден рядом с jar", e);
        }

        GDocsBot bot = new GDocsBot(token, username);

        Runtime.getRuntime().addShutdownHook(new Thread(() -> {
            LOG.info("Завершение работы бота...");
            bot.shutdown();
            LOG.info("Бот остановлен.");
        }, "shutdown-hook"));

        TelegramBotsApi api = new TelegramBotsApi(DefaultBotSession.class);
        api.registerBot(bot);
        LOG.info("Бот запущен: @" + username);
    }

    private static void setupLogging() throws IOException {
        Logger root = Logger.getLogger("");
        root.setLevel(Level.INFO);

        ConsoleHandler console = new ConsoleHandler();
        console.setFormatter(new SimpleFormatter());
        console.setLevel(Level.INFO);
        root.addHandler(console);

        FileHandler file = new FileHandler("bot.log", 10 * 1024 * 1024, 5, true);
        file.setFormatter(new SimpleFormatter());
        file.setLevel(Level.INFO);
        root.addHandler(file);
    }

    enum Google {
        Presentation(1920, 1080,
                "https://docs.google.com/presentation/d/",
                "docs\\.google\\.com/presentation/d/([a-zA-Z0-9_-]+)"),
        Document(1000, 1150,
                "https://docs.google.com/document/d/",
                "docs\\.google\\.com/document/d/([a-zA-Z0-9_-]+)");

        static final int DOC_WIDTH = 820;
        static final int DOC_SCROLL_STEP = 1133;
        static final String PREVIEW = "/preview";
        static final String MOBILE_BASIC = "/mobilebasic";
        static final String HTML_PRESENT = "/htmlpresent";
        static final double SCALE = 2.0;
        static final int MAX_PAGES = 1000;
        static final long MAX_FILE_BYTES = 50L * 1024 * 1024;

        final int width, height;
        final String prefix, regex;

        Google(int w, int h, String p, String r) {
            width = w;
            height = h;
            prefix = p;
            regex = r;
        }
    }

    record Link(Google t, String url) {
    }

    enum Mode {
        OCR, OCR_ONLY_TEXT, HTML
    }

    enum State {
        AWAITING_MODE, AWAITING_URL, PROCESSING
    }

    static class Session {
        volatile Mode mode = null;
        volatile State state = State.AWAITING_MODE;
    }

    static class GDocsBot extends TelegramLongPollingBot {

        private static final Logger LOG = Logger.getLogger(GDocsBot.class.getName());

        private static final int MAX_PARALLEL = 4;

        private final String username;

        private final Map<Long, Session> sessions = Collections.synchronizedMap(
                new LinkedHashMap<>(256, 0.75f, true) {
                    @Override
                    protected boolean removeEldestEntry(Map.Entry<Long, Session> eldest) {
                        return size() > 10_000;
                    }
                });

        private final ThreadPoolExecutor executor = (ThreadPoolExecutor) Executors.newFixedThreadPool(MAX_PARALLEL,
                r -> {
                    Thread t = new Thread(r, "converter-thread");
                    t.setDaemon(false);
                    return t;
                });

        private volatile Path tessDataDir = null;
        private final Object tessInitLock = new Object();

        public GDocsBot(String token, String username) {
            super(token);
            this.username = username;
        }

        @Override
        public String getBotUsername() {
            return username;
        }

        public void shutdown() {
            executor.shutdown();
            try {
                if (!executor.awaitTermination(30, TimeUnit.SECONDS)) {
                    LOG.warning("Принудительное завершение потоков...");
                    executor.shutdownNow();
                }
            } catch (InterruptedException e) {
                executor.shutdownNow();
                Thread.currentThread().interrupt();
            }
        }

        @Override
        public void onUpdateReceived(Update update) {
            if (!update.hasMessage() || !update.getMessage().hasText())
                return;

            long chatId = update.getMessage().getChatId();
            String text = update.getMessage().getText().trim();

            Session session = sessions.computeIfAbsent(chatId, id -> new Session());

            if (text.equals("/start") || text.equals("/reset")) {
                if (session.state == State.PROCESSING) {
                    sendText(chatId, "Обработка уже идёт, подожди...", null);
                    return;
                }
                session.state = State.AWAITING_MODE;
                session.mode = null;
                sendModeKeyboard(chatId, "Привет! Выбери режим обработки:");
                return;
            }

            switch (session.state) {
                case AWAITING_MODE -> handleMode(chatId, text, session);
                case AWAITING_URL -> handleUrl(chatId, text, session);
                case PROCESSING -> sendText(chatId,
                        "Уже обрабатываю твой запрос, подожди...", null);
            }
        }

        private void handleMode(long chatId, String text, Session session) {
            Mode mode = switch (text) {
                case "OCR" -> Mode.OCR;
                case "OCR+OnlyText" -> Mode.OCR_ONLY_TEXT;
                case "HTML" -> Mode.HTML;
                default -> null;
            };

            if (mode == null) {
                sendModeKeyboard(chatId, "Выбери режим с помощью кнопок:");
                return;
            }

            session.mode = mode;
            session.state = State.AWAITING_URL;
            sendText(chatId,
                    "Режим *" + text + "* выбран.\n\nПришли ссылку на Google Slides или Google Docs:",
                    removeKeyboard());
        }

        private void handleUrl(long chatId, String text, Session session) {
            if (session.mode == null) {
                session.state = State.AWAITING_MODE;
                sendModeKeyboard(chatId, "Сначала выбери режим:");
                return;
            }

            if (!text.startsWith("http://") && !text.startsWith("https://")) {
                sendText(chatId, "Не похоже на URL. Пришли ссылку, начинающуюся с https://", null);
                return;
            }

            Link link = parseLink(text, session.mode);
            if (link == null) {
                sendText(chatId,
                        "Ссылка не распознана. Поддерживаются:\n" +
                                "• docs.google.com/presentation/d/...\n" +
                                "• docs.google.com/document/d/...",
                        null);
                return;
            }

            int queued = executor.getQueue().size();
            if (queued > 0) {
                sendText(chatId, "В очереди " + queued +
                        " задач. Твой запрос принят, ожидай...", null);
            } else {
                sendText(chatId, "Начинаю обработку, это может занять несколько минут...", null);
            }

            session.state = State.PROCESSING;

            executor.submit(() -> {
                File result = null;
                try {
                    result = process(link, session.mode, chatId);
                    sendDocument(chatId, result, "Готово!");
                } catch (Exception e) {
                    LOG.log(Level.SEVERE, "Ошибка обработки для chatId=" + chatId, e);
                    sendText(chatId, "Ошибка: " + e.getMessage(), null);
                } finally {
                    if (result != null)
                        result.delete();
                    session.state = State.AWAITING_MODE;
                    session.mode = null;
                    sendModeKeyboard(chatId, "Хочешь обработать ещё? Выбери режим:");
                }
            });
        }

        private File process(Link link, Mode mode, long chatId) throws Exception {
            if (mode == Mode.HTML) {
                File out = File.createTempFile("gdocs_", ".html");
                try (FileWriter fw = new FileWriter(out, StandardCharsets.UTF_8)) {
                    logUser(chatId, "HTML: " + link.url());
                    parseSinglePageHtmlMode(link, fw, chatId);
                }
                return out;
            }

            boolean ocr = (mode == Mode.OCR || mode == Mode.OCR_ONLY_TEXT);
            boolean onlyTxt = (mode == Mode.OCR_ONLY_TEXT);

            ensureChromiumInstalled(chatId);
            if (ocr)
                ensureTessDataReady(chatId);

            File out = File.createTempFile("gdocs_", ".pdf");
            try (PDDocument pdf = new PDDocument();
                    Playwright playwright = Playwright.create();
                    Browser browser = playwright.chromium().launch(
                            new BrowserType.LaunchOptions().setHeadless(true));
                    BrowserContext ctx = browser.newContext(
                            new Browser.NewContextOptions()
                                    .setDeviceScaleFactor(Google.SCALE));
                    Page page = ctx.newPage()) {

                Tesseract localTess = ocr ? buildTesseract() : null;
                parsePage(page, link, pdf, ocr, onlyTxt, chatId, localTess);
                pdf.save(out);
            }
            return out;
        }

        private void ensureTessDataReady(long chatId) throws IOException {
            if (tessDataDir != null)
                return;
            synchronized (tessInitLock) {
                if (tessDataDir != null)
                    return;
                logUser(chatId, "Подготовка tessdata...");
                Path tmp = Files.createTempDirectory("gdocs2pdf");
                Path data = tmp.resolve("tessdata");
                Files.createDirectory(data);
                for (String f : List.of("rus.traineddata", "eng.traineddata", "osd.traineddata")) {
                    InputStream is = getClass().getResourceAsStream("/tessdata/" + f);
                    if (is == null)
                        throw new IOException("Отсутствует: /tessdata/" + f);
                    Files.copy(is, data.resolve(f));
                }
                tessDataDir = data;
                logUser(chatId, "tessdata готова");
            }
        }

        private Tesseract buildTesseract() {
            Tesseract t = new Tesseract();
            t.setDatapath(tessDataDir.toAbsolutePath().toString());
            t.setLanguage("rus+eng");
            t.setPageSegMode(1);
            t.setOcrEngineMode(1);
            t.setVariable("user_defined_dpi", String.valueOf((int) (96 * Google.SCALE)));
            return t;
        }

        private void parseSinglePageHtmlMode(Link link, FileWriter out, long chatId)
                throws IOException {
            Document doc = Jsoup.connect(link.url())
                    .userAgent("Mozilla/5.0 (Linux; Android 12; Pixel 6) " +
                            "AppleWebKit/537.36 (KHTML, like Gecko) " +
                            "Chrome/120.0.0.0 Mobile Safari/537.36")
                    .timeout(30_000)
                    .get();

            String title = doc.title().trim().isEmpty() ? "Документ" : doc.title().trim();
            out.append("<!DOCTYPE html><html><head>")
                    .append("<meta charset='UTF-8'>")
                    .append("<title>").append(escapeHtml(title)).append("</title>")
                    .append("</head><body>\n");

            renderNode(doc.body(), out);
            out.append("</body></html>");
            logUser(chatId, "HTML собран: " + title);
        }

        private void renderNode(Node node, FileWriter out) throws IOException {
            if (node instanceof TextNode tn) {
                String t = tn.text();
                if (!t.isBlank())
                    out.append(escapeHtml(t));
                return;
            }
            if (!(node instanceof Element el))
                return;

            String tag = el.tagName().toLowerCase();
            switch (tag) {
                case "h1", "h2", "h3", "h4", "h5", "h6" -> {
                    out.append("<").append(tag).append(">");
                    renderInlineContent(el, out);
                    out.append("</").append(tag).append(">\n");
                }
                case "p" -> {
                    out.append("<p>");
                    renderInlineContent(el, out);
                    out.append("</p>\n");
                }
                case "div", "section", "article", "main", "aside", "header", "footer", "nav", "form" -> {
                    if (!el.ownText().isBlank()) {
                        out.append("<p>");
                        renderInlineContent(el, out);
                        out.append("</p>\n");
                    } else {
                        for (Node child : el.childNodes())
                            renderNode(child, out);
                    }
                }
                case "ul" -> {
                    out.append("<ul>\n");
                    for (Node child : el.childNodes())
                        renderNode(child, out);
                    out.append("</ul>\n");
                }
                case "ol" -> {
                    out.append("<ol>\n");
                    for (Node child : el.childNodes())
                        renderNode(child, out);
                    out.append("</ol>\n");
                }
                case "li" -> {
                    out.append("<li>");
                    renderInlineContent(el, out);
                    for (Element nested : el.select("> ul, > ol"))
                        renderNode(nested, out);
                    out.append("</li>\n");
                }
                case "table" -> renderTable(el, out);
                case "br" -> out.append("<br>\n");
                case "hr" -> out.append("<hr>\n");
                case "pre", "code" -> out.append("<").append(tag).append(">")
                        .append(escapeHtml(el.wholeText()))
                        .append("</").append(tag).append(">\n");
                case "blockquote" -> {
                    out.append("<blockquote>");
                    renderInlineContent(el, out);
                    out.append("</blockquote>\n");
                }
                default -> {
                    for (Node child : el.childNodes())
                        renderNode(child, out);
                }
            }
        }

        private void renderInlineContent(Element parent, FileWriter out) throws IOException {
            for (Node node : parent.childNodes()) {
                if (node instanceof TextNode tn) {
                    String t = tn.text();
                    if (!t.isEmpty())
                        out.append(escapeHtml(t));
                } else if (node instanceof Element el) {
                    String tag = el.tagName().toLowerCase();
                    switch (tag) {
                        case "b", "strong" -> {
                            out.append("<b>");
                            renderInlineContent(el, out);
                            out.append("</b>");
                        }
                        case "i", "em" -> {
                            out.append("<i>");
                            renderInlineContent(el, out);
                            out.append("</i>");
                        }
                        case "u" -> {
                            out.append("<u>");
                            renderInlineContent(el, out);
                            out.append("</u>");
                        }
                        case "s", "strike", "del" -> {
                            out.append("<s>");
                            renderInlineContent(el, out);
                            out.append("</s>");
                        }
                        case "sup" -> {
                            out.append("<sup>");
                            renderInlineContent(el, out);
                            out.append("</sup>");
                        }
                        case "sub" -> {
                            out.append("<sub>");
                            renderInlineContent(el, out);
                            out.append("</sub>");
                        }
                        case "a" -> {
                            String href = el.attr("abs:href");
                            if (href.isEmpty())
                                href = el.attr("href");
                            out.append("<a href='").append(escapeHtml(href)).append("'>");
                            renderInlineContent(el, out);
                            out.append("</a>");
                        }
                        case "br" -> out.append("<br>\n");
                        case "span", "font", "label" -> {
                            String style = el.attr("style");
                            boolean bold = isCssBold(style);
                            boolean italic = isCssItalic(style);
                            if (bold)
                                out.append("<b>");
                            if (italic)
                                out.append("<i>");
                            renderInlineContent(el, out);
                            if (italic)
                                out.append("</i>");
                            if (bold)
                                out.append("</b>");
                        }
                        case "code" -> out.append("<code>").append(escapeHtml(el.text())).append("</code>");
                        default -> renderInlineContent(el, out);
                    }
                }
            }
        }

        private void renderTable(Element table, FileWriter out) throws IOException {
            out.append("<table border='1'>\n");
            for (Element row : table.select("tr")) {
                if (!row.parents().first().closest("table").equals(table))
                    continue;
                out.append("  <tr>");
                for (Element cell : row.select("td, th")) {
                    StringBuilder attrs = new StringBuilder();
                    String colspan = cell.attr("colspan");
                    String rowspan = cell.attr("rowspan");
                    if (!colspan.isEmpty() && !colspan.equals("1"))
                        attrs.append(" colspan='").append(escapeHtml(colspan)).append("'");
                    if (!rowspan.isEmpty() && !rowspan.equals("1"))
                        attrs.append(" rowspan='").append(escapeHtml(rowspan)).append("'");
                    String ct = cell.tagName();
                    out.append("<").append(ct).append(attrs).append(">");
                    if (cell.text().trim().isEmpty())
                        out.append("&nbsp;");
                    else
                        renderInlineContent(cell, out);
                    out.append("</").append(ct).append(">");
                }
                out.append("</tr>\n");
            }
            out.append("</table>\n");
        }

        private void parsePage(Page page, Link link, PDDocument pdf,
                boolean ocr, boolean txt,
                long chatId, Tesseract tess) throws IOException {
            page.setViewportSize(link.t().width, link.t().height);

            PDFont font = null;
            if (ocr) {
                InputStream is = getClass().getResourceAsStream("/fonts/LiberationSans-Regular.ttf");
                if (is == null)
                    throw new IOException("Отсутствует: /fonts/LiberationSans-Regular.ttf");
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
                            throw new IOException("Ошибка декодирования скриншота");
                        int cropWidth = Math.min((int) (Google.DOC_WIDTH * Google.SCALE), original.getWidth());
                        int x = Math.max(0, (original.getWidth() - cropWidth) / 2);
                        BufferedImage cropped = original.getSubimage(x, 0, cropWidth, original.getHeight());
                        ByteArrayOutputStream baos = new ByteArrayOutputStream();
                        if (!ImageIO.write(cropped, "png", baos))
                            throw new IOException("PNG writer недоступен");
                        shot = baos.toByteArray();
                    }

                    if (Arrays.equals(prev, shot)) {
                        logUser(chatId, "Страниц всего: " + saved);
                        return;
                    }

                    try {
                        addPage(pdf, shot, num, font, txt, chatId, tess);
                        saved++;
                        if (saved == 1 || saved % 10 == 0)
                            logUser(chatId, "Обработано страниц: " + saved);
                    } catch (Exception e) {
                        LOG.log(Level.WARNING, "Страница " + num + " пропущена", e);
                        logUser(chatId, "Страница " + num + " пропущена: " + e.getMessage());
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
                    LOG.log(Level.WARNING, "Страница " + num + " недоступна", e);
                    logUser(chatId, "Страница " + num + " недоступна: " + e.getMessage());
                    if (link.t() == Google.Presentation && prev != null) {
                        logUser(chatId, "Страниц всего: " + saved);
                        return;
                    }
                }
            }
            logUser(chatId, "ВНИМАНИЕ: достигнут лимит " + Google.MAX_PAGES + " страниц");
        }

        private void addPage(PDDocument pdf, byte[] shot, int num,
                PDFont font, boolean txt,
                long chatId, Tesseract tess) throws IOException {
            BufferedImage image = ImageIO.read(new ByteArrayInputStream(shot));
            if (image == null)
                throw new IOException("Ошибка декодирования скриншота на стр. " + num);

            PDImageXObject pdImg = PDImageXObject.createFromByteArray(pdf, shot, "p" + num);
            float pageWidth = pdImg.getWidth() / (float) Google.SCALE;
            float pageHeight = pdImg.getHeight() / (float) Google.SCALE;
            PDPage pdfPage = new PDPage(new PDRectangle(pageWidth, pageHeight));
            pdf.addPage(pdfPage);

            try (PDPageContentStream stream = new PDPageContentStream(pdf, pdfPage)) {
                if (!txt)
                    stream.drawImage(pdImg, 0, 0, pageWidth, pageHeight);
                if (font == null || tess == null)
                    return;

                List<Word> lines;
                try {
                    lines = tess.getWords(image,
                            txt ? ITessAPI.TessPageIteratorLevel.RIL_TEXTLINE
                                    : ITessAPI.TessPageIteratorLevel.RIL_WORD);
                } catch (Exception e) {
                    LOG.log(Level.WARNING, "OCR ошибка на стр. " + num, e);
                    return;
                }
                if (lines == null || lines.isEmpty())
                    return;

                stream.setRenderingMode(txt ? RenderingMode.FILL : RenderingMode.NEITHER);

                float scaleX = pageWidth / image.getWidth();
                float scaleY = pageHeight / image.getHeight();
                List<Float> recentSizes = new ArrayList<>();

                for (Word line : lines) {
                    String text = line.getText();
                    if (text == null)
                        continue;
                    text = text.trim();
                    if (text.isEmpty())
                        continue;

                    Rectangle bbox = line.getBoundingBox();
                    float pdfX = bbox.x * scaleX;
                    float fontSize, pdfY, horizScaling;

                    if (!txt) {
                        float boxWidth = bbox.width * scaleX;
                        float boxHeight = bbox.height * scaleY;
                        fontSize = Math.max(boxHeight * 0.85f, 4f);
                        pdfY = pageHeight - (bbox.y + bbox.height) * scaleY + boxHeight * 0.15f;
                        float textWidth;
                        try {
                            textWidth = font.getStringWidth(text) / 1000f * fontSize;
                        } catch (Exception e) {
                            continue;
                        }
                        if (textWidth <= 0)
                            continue;
                        horizScaling = (boxWidth / textWidth) * 100f;
                    } else {
                        pdfY = pageHeight - (bbox.y + bbox.height) * scaleY;
                        recentSizes.add(bbox.height * scaleY * 0.85f);
                        if (recentSizes.size() > 5)
                            recentSizes.removeFirst();
                        float sum = 0f;
                        for (Float s : recentSizes)
                            sum += s;
                        fontSize = Math.max(sum / recentSizes.size(), 8f);
                        horizScaling = 100f;
                    }

                    stream.setFont(font, fontSize);
                    stream.setHorizontalScaling(horizScaling);
                    stream.beginText();
                    stream.newLineAtOffset(pdfX, pdfY);
                    try {
                        stream.showText(text);
                    } catch (Exception e) {
                        LOG.warning("OCR showText ошибка стр. " + num + ": " + e.getMessage());
                    } finally {
                        stream.endText();
                    }
                }
                stream.setHorizontalScaling(100f);
            }
        }

        private Link parseLink(String rawUrl, Mode mode) {
            boolean html = (mode == Mode.HTML);

            Matcher slides = Pattern.compile(Google.Presentation.regex).matcher(rawUrl);
            if (slides.find()) {
                return new Link(Google.Presentation,
                        Google.Presentation.prefix + slides.group(1) +
                                (html ? Google.HTML_PRESENT : Google.PREVIEW));
            }
            Matcher docs = Pattern.compile(Google.Document.regex).matcher(rawUrl);
            if (docs.find()) {
                return new Link(Google.Document,
                        Google.Document.prefix + docs.group(1) +
                                (html ? Google.MOBILE_BASIC : Google.PREVIEW));
            }
            return null;
        }

        private void ensureChromiumInstalled(long chatId) throws IOException, InterruptedException {
            if (Files.exists(getCachePath()))
                return;
            logUser(chatId, "Устанавливаю Chromium (первый запуск)...");
            ProcessBuilder pb = new ProcessBuilder(
                    "java", "-cp", System.getProperty("java.class.path"),
                    "com.microsoft.playwright.CLI", "install", "chromium");
            pb.inheritIO();
            int code = pb.start().waitFor();
            if (code != 0)
                throw new IOException("Chromium: код выхода " + code);
            logUser(chatId, "Chromium установлен");
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

        private static boolean isCssBold(String style) {
            if (style == null || style.isEmpty())
                return false;
            if (style.matches("(?i).*font-weight\\s*:\\s*(bold|bolder).*"))
                return true;
            Matcher m = Pattern.compile("(?i)font-weight\\s*:\\s*(\\d+)").matcher(style);
            return m.find() && Integer.parseInt(m.group(1)) >= 700;
        }

        private static boolean isCssItalic(String style) {
            if (style == null || style.isEmpty())
                return false;
            return style.matches("(?i).*font-style\\s*:\\s*(italic|oblique).*");
        }

        private static String escapeHtml(String text) {
            if (text == null)
                return "";
            return text.replace("&", "&amp;").replace("<", "&lt;")
                    .replace(">", "&gt;").replace("\"", "&quot;").replace("'", "&#39;");
        }

        private void logUser(long chatId, String msg) {
            LOG.info("[" + chatId + "] " + msg);
            sendText(chatId, msg, null);
        }

        private void sendModeKeyboard(long chatId, String text) {
            KeyboardRow row1 = new KeyboardRow();
            row1.add("OCR");
            KeyboardRow row2 = new KeyboardRow();
            row2.add("OCR+OnlyText");
            KeyboardRow row3 = new KeyboardRow();
            row3.add("HTML");

            sendText(chatId, text, ReplyKeyboardMarkup.builder()
                    .keyboard(List.of(row1, row2, row3))
                    .resizeKeyboard(true)
                    .oneTimeKeyboard(false)
                    .build());
        }

        private void sendText(long chatId, String text,
                org.telegram.telegrambots.meta.api.objects.replykeyboard.ReplyKeyboard keyboard) {
            try {
                execute(SendMessage.builder()
                        .chatId(String.valueOf(chatId))
                        .text(text)
                        .parseMode("Markdown")
                        .replyMarkup(keyboard)
                        .build());
            } catch (Exception e) {
                LOG.log(Level.WARNING, "sendText ошибка chatId=" + chatId, e);
            }
        }

        private void sendDocument(long chatId, File file, String caption) {
            if (file.length() > Google.MAX_FILE_BYTES) {
                long mb = file.length() / 1024 / 1024;
                sendText(chatId,
                        "Файл слишком большой (" + mb + " MB). Лимит Telegram — 50 MB.\n" +
                                "Попробуй режим HTML или уменьши документ.",
                        null);
                return;
            }
            try {
                execute(SendDocument.builder()
                        .chatId(String.valueOf(chatId))
                        .document(new InputFile(file))
                        .caption(caption)
                        .build());
            } catch (Exception e) {
                LOG.log(Level.SEVERE, "sendDocument ошибка chatId=" + chatId, e);
                sendText(chatId, "Не удалось отправить файл: " + e.getMessage(), null);
            }
        }

        private ReplyKeyboardRemove removeKeyboard() {
            return ReplyKeyboardRemove.builder().removeKeyboard(true).build();
        }
    }
}
