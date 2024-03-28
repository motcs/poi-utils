package com.motcs.utils.poi;

import com.motcs.utils.poi.entity.ImageUrl;
import com.motcs.utils.poi.enums.ChineseFont;
import com.motcs.utils.poi.enums.Color;
import com.motcs.utils.poi.enums.Footer;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFldChar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.net.URL;

import static org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_JPEG;

/**
 * HTW html to word
 * html富文本内容转word
 *
 * @author <a href="https://github.com/motcs">motcs</a>
 * @since 2024-03-28 星期四
 */
@Log4j2
public class HTWConverter {

    /**
     * 设置文档的页边距 单位为dax,使用厘米时请使用方法：{@link #pageMarConvert(double)} 转换
     * <p></p>
     * 示例 pageMarConvert(2.0) 转换结果：2.0厘米
     * <p></p>
     *
     * @param document 文档
     * @param left     左边距
     * @param right    右边距
     * @param top      上边距
     * @param bottom   下边距
     */
    public static void pageMar(XWPFDocument document, long left, long right, long top, long bottom) {
        CTPageMar pageMar = document.getDocument().getBody().addNewSectPr().addNewPgMar();
        pageMar.setLeft(BigInteger.valueOf(left));
        pageMar.setRight(BigInteger.valueOf(right));
        pageMar.setTop(BigInteger.valueOf(top));
        pageMar.setBottom(BigInteger.valueOf(bottom));
    }

    /**
     * 页边距单位换算，换算后单位厘米
     *
     * @param value 边距
     * @return 返回换算后的数值
     */
    public static long pageMarConvert(double value) {
        return (long) ((value / 2.54) * 1440);
    }

    /**
     * 解析数据
     *
     * @param document    插入的文档
     * @param htmlContent 需要解析的内容
     */
    public static void processHtmlContent(XWPFDocument document, String htmlContent) {
        Elements elements = processHtmlContent(htmlContent, "p, img, h1");
        for (Element element : elements) {
            if (element.tagName().equals("p")) {
                addParagraphToDocument(document, null, 0, false, element.text());
            } else if (element.tagName().equals("img")) {
                addImageToDocument(document, ImageUrl.builder().url(element.attr("src")).build());
            } else if (element.tagName().equals("h1")) {
                addParagraphHToDocument(document, null, 0, false, element.text());
            }
        }
    }

    /**
     * 读取数据
     * tips： 主要用来处理需要单独定义的
     * 使用具体见方法
     * <p>
     * {@link #processHtmlContent(XWPFDocument, String)}
     *
     * @param htmlContent 需解析的正文内容
     * @param cssQuery    解析的标签如：p, img, h1 等html标签
     * @return 返回解析后的分段内容
     */
    public static Elements processHtmlContent(String htmlContent, String cssQuery) {
        return Jsoup.parse(htmlContent).select(cssQuery);
    }

    /**
     * 插入正文，字体 仿宋GB2312 字号： 三号
     *
     * @param document   文档
     * @param fontFamily 字体类型
     * @param fontSize   字体大小
     * @param text       正文内容
     */
    public static void addParagraphToDocument(XWPFDocument document, String fontFamily,
                                              int fontSize, boolean bold, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.setFontSize(fontSize > 0 ? fontSize : 16);
        run.setFontFamily(fontFamily == null ? ChineseFont.FANG_SONG_GB2312.getFontName() : fontFamily);
        run.setBold(bold);
        Double fontSizeAsDouble = run.getFontSizeAsDouble();
        paragraph.setIndentationFirstLine(fontSizeAsDouble == null
                || fontSizeAsDouble <= 0 ? 420 : fontSizeAsDouble.intValue() * 2 * 20);
    }

    /**
     * 设置标题
     *
     * @param document 文档
     * @param text     标题
     */
    public static void addParagraphHToDocument(XWPFDocument document, String fontFamily,
                                               int fontSize, boolean bold, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.setFontSize(fontSize > 0 ? fontSize : 16);
        run.setFontFamily(fontFamily == null ? ChineseFont.HEI_TI.getFontName() : fontFamily);
        run.setBold(bold);
    }

    /**
     * 文档增加图片
     *
     * @param document 文档
     * @param image1   图片信息
     */
    public static void addImageToDocument(XWPFDocument document, ImageUrl image1) {
        String imageUrl = image1.getUrl();
        try {
            BufferedImage image = ImageIO.read(new URL(imageUrl));
            String fileExtension = imageUrl.substring(imageUrl.lastIndexOf('.') + 1);
            log.info("图片后缀: {}", fileExtension);
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            ImageIO.write(image, fileExtension, os);

            try (InputStream is = new ByteArrayInputStream(os.toByteArray())) {
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.addPicture(is, PICTURE_TYPE_JPEG, "image." + fileExtension,
                        (int) Math.rint(image1.getWidth() * Units.EMU_PER_CENTIMETER),
                        (int) Math.rint(image1.getHeight() * Units.EMU_PER_CENTIMETER));
            } catch (InvalidFormatException e) {
                log.info("图片路径：{}，解析失败：{}", imageUrl, e.getMessage());
                addParagraphToDocument(document, null, 0, false, "");
                addParagraphToDocument(document, null, 0, false, image1.getErrMsgD());
                addParagraphToDocument(document, null, 0, false, "");
            }
        } catch (Exception e) {
            log.info("下载图片失败:{}", e.getMessage());
            addParagraphToDocument(document, null, 0, false, "");
            addParagraphToDocument(document, null, 0, false, image1.getErrMsgA());
            addParagraphToDocument(document, null, 0, false, "");
        }
    }

    /**
     * 设置页脚，默认居中显示
     *
     * @param document   文档
     * @param fontFamily 字体
     * @param fontSize   字体大小
     * @param color      字体颜色
     * @param bold       是否加粗
     * @param prefix     页码前缀  如 "第 1 页"  的 "第"
     * @param suffix     页码后缀  如 "第 1 页"  的 "页"
     */
    public static void setFooter(XWPFDocument document, String fontFamily, ParagraphAlignment alignment,
                                 int fontSize, String color, boolean bold, String prefix, String suffix) {
        XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);
        //创建一个新的XWPFFooter对象，HeaderFooterType.DEFAULT表示所有页
        XWPFParagraph paragraph = footer.createParagraph();
        //创建新的XWPFParagraph对象
        paragraph.setAlignment(alignment == null ? ParagraphAlignment.CENTER : alignment);
        appendPageNumber(paragraph, fontFamily, fontSize, color, bold, prefix, suffix);
    }

    /**
     * 段落后拼接页码
     *
     * @param paragraph  段落
     * @param fontFamily 字体
     * @param fontSize   字体大小
     * @param color      字体颜色
     * @param bold       是否加粗
     * @param prefix     页码前缀  如 "第 1 页"  的 "第"
     * @param suffix     页码后缀  如 "第 1 页"  的 "页"
     */
    public static void appendPageNumber(XWPFParagraph paragraph, String fontFamily, int fontSize,
                                        String color, boolean bold, String prefix, String suffix) {
        if (color == null || color.isEmpty() || color.trim().isEmpty()) {
            // 默认黑色
            color = Color.BLACK.getValue();
        }
        setStyle(paragraph.createRun(), fontFamily, fontSize, bold, prefix, color);
        CTFldChar fldChar = paragraph.createRun().getCTR().addNewFldChar();
        fldChar.setFldCharType(STFldCharType.Enum.forString(Footer.BEGIN.getValue()));
        XWPFRun numberRun = paragraph.createRun();
        CTText ctText = numberRun.getCTR().addNewInstrText();
        ctText.setStringValue(Footer.PAGE.getValue());
        ctText.setSpace(SpaceAttribute.Space.Enum.forString(Footer.PRESERVE.getValue()));
        setStyle(numberRun, fontFamily, fontSize, bold, null, color);
        fldChar = paragraph.createRun().getCTR().addNewFldChar();
        fldChar.setFldCharType(STFldCharType.Enum.forString(Footer.END.getValue()));
        setStyle(paragraph.createRun(), fontFamily, fontSize, bold, suffix, color);
    }

    /**
     * 设置样式
     *
     * @param fontFamily 字体
     * @param fontSize   字体大小
     * @param bold       是否加粗
     * @param text       设置的文本
     * @param color      字体颜色
     */
    public static void setStyle(XWPFRun run, String fontFamily, int fontSize,
                                boolean bold, String text, String color) {
        run.setBold(bold);
        run.setFontFamily(fontFamily);
        run.setFontSize(fontSize);
        if (text != null && !text.isEmpty()) {
            run.setText(text);
        } //颜色默认黑色
        run.setColor(color != null ? color : Color.BLACK.getValue());
    }

}
