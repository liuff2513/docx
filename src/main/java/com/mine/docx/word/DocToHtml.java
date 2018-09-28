package com.mine.docx.word;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Administrator on 2018/9/28.
 */
public class DocToHtml {
    private static final String encoding = "UTF-8";

    public static String convert2Html(String wordPath)
            throws FileNotFoundException, TransformerException, IOException,
            ParserConfigurationException {
        if (wordPath == null || "".equals(wordPath)) return "";
        File file = new File(wordPath);
        if (file.exists() && file.isFile())
            return convert2Html(new FileInputStream(file));
        else
            return "";
    }

    public static String convert2Html(String wordPath, String context)
            throws FileNotFoundException, TransformerException, IOException,
            ParserConfigurationException {
        if (wordPath == null || "".equals(wordPath)) return "";
        File file = new File(wordPath);
        if (file.exists() && file.isFile())
            return convert2Html(new FileInputStream(file), context);
        else
            return "";
    }

    public static String convert2Html(InputStream is)
            throws TransformerException, IOException,
            ParserConfigurationException {
        return convert2Html(is, "");
    }

    public static String convert2Html(InputStream is, HttpServletRequest req) throws TransformerException, IOException, ParserConfigurationException {
        return convert2Html(is, req.getContextPath());
    }

    public static String convert2Html(InputStream is, final String context) throws IOException, ParserConfigurationException, TransformerException {
        HWPFDocument wordDocument = new HWPFDocument(is);
        WordToHtmlConverter converter = new WordToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder()
                        .newDocument());

        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSSS");
        final String prefix = sdf.format(new Date());
        final Map<Object, String> suffixMap = new HashMap<Object, String>();

        converter.setPicturesManager(new PicturesManager() {
            public String savePicture(byte[] content, PictureType pictureType,
                                      String suggestedName, float widthInches, float heightInches) {
                String prefixContext = context.replace("\\", "").replace("/", "");
                prefixContext = StringUtils.isNotBlank(prefixContext) ? "/" + prefixContext + "/" : prefixContext;
                suffixMap.put(new String(content).replace(" ", "").length(), suggestedName);

                return prefixContext
                        + UeConstants.VIEW_IMAGE_PATH + "/" + UeConstants.UEDITOR_PATH
                        + "/" + UeConstants.UEDITOR_IMAGE_PATH + "/"
                        + prefix + "_"
                        + suggestedName;
            }
        });
        converter.processDocument(wordDocument);

        List<Picture> pics = wordDocument.getPicturesTable().getAllPictures();
        if (pics != null) {
            for (Picture pic : pics) {
                try {
                    pic.writeImageContent(new FileOutputStream(
                            UeConstants.IMAGE_PATH
                                    + "/" + prefix + "_" + suffixMap.get(new String(pic.getContent()).replace(" ", "").length())));
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                }
            }
        }

        StringWriter writer = new StringWriter();

        Transformer serializer = TransformerFactory.newInstance().newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, encoding);
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(
                new DOMSource(converter.getDocument()),
                new StreamResult(writer));
        writer.close();
        return writer.toString();
    }
}
