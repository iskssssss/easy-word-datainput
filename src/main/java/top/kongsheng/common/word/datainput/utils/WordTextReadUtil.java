package top.kongsheng.common.word.datainput.utils;

import cn.hutool.core.util.StrUtil;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import top.kongsheng.common.word.datainput.core.CloneByteArrayInputStream;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.math.RoundingMode;
import java.util.*;

/**
 * word文本读取
 *
 * @author 孔胜
 * @date 2023/1/3 16:25
 */
public class WordTextReadUtil {

    public static String read(CloneByteArrayInputStream cloneByteArrayInputStream) throws Exception {
        FileMagic fm = FileMagic.valueOf(cloneByteArrayInputStream.get());
        switch (fm) {
            case OLE2:
                System.out.println("OLE2");
                return WordTextReadUtil.OLE2_read(cloneByteArrayInputStream.get());
            case OOXML:
                System.out.println("OOXML");
                return WordTextReadUtil.OOXML_read(cloneByteArrayInputStream.get());
            case XML:
            default:
                return null;
        }
    }

    public static String XML_read(InputStream inputStream) throws IOException {
        XWPFDocument document = new XWPFDocument(inputStream);
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        List<String> paragraphList = new LinkedList<>();
        for (XWPFParagraph paragraph : paragraphs) {
            StringBuilder sb = new StringBuilder();
            for (XWPFRun run : paragraph.getRuns()) {
                sb.append(run.text());
            }
            String result = sb.toString();
            if (StrUtil.isEmpty(result)) {
                continue;
            }
            paragraphList.add(result);
        }
        return String.join("\n", paragraphList);
    }

    public static String OOXML_read(InputStream inputStream) throws Exception {
        return String.join("\n", OOXMLReadList(inputStream, false));
    }

    public static List<String> OOXMLReadList(InputStream inputStream, boolean close) throws Exception {
        if (close) {
            try (XWPFDocument document = new XWPFDocument(inputStream)) {
                OOXMLWordRead ooxmlWordRead = new OOXMLWordRead(document);
                return ooxmlWordRead.read();
            }
        }
        XWPFDocument document = new XWPFDocument(inputStream);
        OOXMLWordRead ooxmlWordRead = new OOXMLWordRead(document);
        return ooxmlWordRead.read();
    }

    static final class OOXMLWordRead {
        private Map<BigInteger, Integer> numMap;
        private final XWPFDocument document;

        public OOXMLWordRead(XWPFDocument document) {
            this.document = document;
            numMap = new HashMap<>();
        }

        public List<String> read() {
            List<String> paragraphList = this.ooxmlReadList(document);
            numMap.clear();
            numMap = null;
            return paragraphList;
        }

        private List<String> ooxmlReadList(XWPFDocument document) {
            Iterator<IBodyElement> bodyElementsIterator = document.getBodyElementsIterator();
            List<String> paragraphList = new LinkedList<>();
            while (bodyElementsIterator.hasNext()) {
                IBodyElement bodyElement = bodyElementsIterator.next();
                StringBuilder sb;
                if (bodyElement instanceof XWPFParagraph) {
                    sb = new StringBuilder();
                    XWPFParagraph paragraph = ((XWPFParagraph) bodyElement);
                    boolean isNum = paragraph.getNumID() != null;
                    if (isNum) {
                        String numStr = createNumStr(paragraph);
                        sb.append(numStr);
                    }
                    List<XWPFRun> runs = paragraph.getRuns();
                    if (runs.isEmpty() && !isNum) {
                        continue;
                    }
                    runs.stream().map(XWPFRun::text).forEach(sb::append);
                } else {
                    continue;
                }
                String result = sb.toString();
                if (StrUtil.isEmpty(result)) {
                    continue;
                }
                paragraphList.add(result);
            }
            return paragraphList;
        }

        private String createNumStr(XWPFParagraph paragraph) {
            BigInteger numId = paragraph.getNumID();
            int num = numMap.computeIfAbsent(numId, key -> 1);
            numMap.put(numId, num + 1);
            String numFmt = paragraph.getNumFmt();
            String numLevelText = paragraph.getNumLevelText();
            switch (numFmt) {
                case "japaneseCounting":
                case "chineseCountingThousand":
                    String number = CharUtil.toUpperCaseNumber(num);
                    return numLevelText.replace("%1", number);
                case "decimal":
                    return numLevelText.replace("%1", num + "");
                case "lowerLetter":
                    return numLevelText.replace("%1", letter(num, LOWER_LETTER_START));
                case "upperLetter":
                    return numLevelText.replace("%1", letter(num, UPPER_LETTER_START));
                default:
            }
            return "";
        }

        private String letter(int num, int letterStart) {
            int count = 1;
            if (num > LETTER_COUNT) {
                count = new BigDecimal(num / ((float) LETTER_COUNT)).setScale(0, RoundingMode.UP).intValue();
                num = num % LETTER_COUNT;
            }
            int upperLetterNum = letterStart + ((num) - 1);
            String value = String.valueOf(((char) upperLetterNum));
            int maxCount = Math.max(0, count);
            StringBuilder stringBuilder = new StringBuilder();
            for (int i = 0; i < maxCount; i++) {
                stringBuilder.append(value);
            }
            return stringBuilder.toString();
        }

        private static final int LOWER_LETTER_START = 97;
        private static final int LOWER_LETTER_END = 122;
        private static final int UPPER_LETTER_START = 65;
        private static final int UPPER_LETTER_END = 90;

        private static final int LETTER_COUNT = 26;
    }

    private static String getNumberingInfo(XWPFDocument document, BigInteger numId) {
        XWPFNumbering numbering = document.getNumbering();
        if (numbering != null) {
            List<XWPFAbstractNum> abstractNums = numbering.getAbstractNums();
            for (XWPFAbstractNum abstractNum : abstractNums) {
                BigInteger abstractNumId = abstractNum.getAbstractNum().getAbstractNumId();
                if (abstractNumId.equals(numId)) {
                    final List<CTLvl> lvlList = abstractNum.getAbstractNum().getLvlList();
                    for (CTLvl ctLvl : lvlList) {
                        System.out.println("start：" + ctLvl.getStart().getVal());
                        System.out.println("lvlText：" + ctLvl.getLvlText().getVal());
                    }
                    return lvlList.toString();
                }
            }
        }
        return null;
    }


    public static String OLE2_read(InputStream inputStream) throws Exception {
//        try {
//            HWPFDocument wordDocument = new HWPFDocument(inputStream);
//            WordToHtmlConverter converter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
//            //对HWPFDocument进行转换
//            converter.processDocument(wordDocument);
//            Transformer transformer = TransformerFactory.newInstance().newTransformer();
//            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
//            transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
//            //是否添加空格
//            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
//            transformer.setOutputProperty(OutputKeys.METHOD, "html");
//            transformer.transform(
//                    new DOMSource(converter.getDocument()),
//                    new StreamResult(outputStream));
//            String result = IoUtil.toStr(outputStream, StandardCharsets.UTF_8);
//            return result;
//        } catch (Exception e) {
        WordExtractor wordExtractor = new WordExtractor(inputStream);
        return wordExtractor.getText();
//        }
        /*POIFSFileSystem pfs = new POIFSFileSystem(inputStream);
        DirectoryNode root = pfs.getRoot();
        Spliterator<Entry> spliterator = root.spliterator();
        StringBuilder resultStr = new StringBuilder();
        spliterator.forEachRemaining(entry -> {
            if (!(entry instanceof DocumentNode)) {
                return;
            }
            DocumentNode documentNode = (DocumentNode) entry;
            POIFSDocument poifsDocument = new POIFSDocument(documentNode);
            int size = poifsDocument.getSize();
            Iterator<ByteBuffer> iterator = poifsDocument.iterator();
            ByteArrayOutputStream os1 = new ByteArrayOutputStream(size);
            int readSize = 0;
            while (iterator.hasNext()) {
                ByteBuffer buffer = iterator.next();
                int limit = buffer.limit();
                int position = buffer.position();
                int read = limit - position;
                if (readSize > size) {
                    continue;
                }
                if (readSize + read > size) {
                    read = size - readSize;
                }
                readSize += read;
                byte[] dst = new byte[read];
                buffer.get(dst);
                try {
                    os1.write(dst);
                } catch (IOException ioException) {
                    ioException.printStackTrace();
                }
            }
            String s1 = IoUtil.toStr(os1, charset);
            resultStr.append(s1);
        });
        return resultStr.toString();*/
    }
}
