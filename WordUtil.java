package com.longshine.superapp.nucleicresult.util;

import com.aliyun.oss.OSSClient;
import com.aliyun.oss.model.GetObjectRequest;
import com.aliyun.oss.model.OSSObject;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.*;

/**
 * @Author Chen ZhiZhe
 * @Create 2021.08.13 13:40
 * @Description word模板自动填充数据
 */
@Component
public class WordUtil {

    @Value("${oss.access-key}")
    public String accessKeyId;

    @Value("${oss.accessKeySecret}")
    public String accessKeySecret;

    @Value("${oss.endpoint}")
    public String endpoint;

    @Value("${oss.bucketName.resource}")
    public String bucketName;

    /**
     * 根据指定的参数值、模板，生成 word 文档
     * 注意：其它模板需要根据情况进行调整
     *
     * @param param    变量集合
     * @param template 模板路径
     */
    public static XWPFDocument generateWord(Map<String, Object> param, String template) {
        XWPFDocument doc = null;
        try {
            OPCPackage pack = POIXMLDocument.openPackage(template);//通过路径获取word模板
            doc = new XWPFDocument(pack);
            if (param != null && param.size() > 0) {
                // 处理段落
                List<XWPFParagraph> paragraphList = doc.getParagraphs();
                processParagraphs(paragraphList, param, doc);
                // 处理表格
                Iterator<XWPFTable> it = doc.getTablesIterator();
                while (it.hasNext()) {

                    XWPFTable table = it.next();
                    List<XWPFTableRow> rows = table.getRows();
                    for (XWPFTableRow row : rows) {

                        List<XWPFTableCell> cells = row.getTableCells();
                        for (XWPFTableCell cell : cells) {

                            List<XWPFParagraph> paragraphListTable = cell.getParagraphs();
                            processParagraphs(paragraphListTable, param, doc);
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return doc;
    }

    /**
     * 处理段落
     */
    public static void processParagraphs(List<XWPFParagraph> paragraphList, Map<String, Object> param, XWPFDocument doc) throws InvalidFormatException, FileNotFoundException {
        if (paragraphList != null && paragraphList.size() > 0) {

            for (XWPFParagraph paragraph : paragraphList) {

                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {

                    String text = run.getText(0);
                    if (text != null) {

                        boolean isSetText = false;
                        for (Map.Entry<String, Object> entry : param.entrySet()) {

                            String key = "${" + entry.getKey() + "}";
                            if (text.contains(key)) {

                                isSetText = true;
                                Object value = entry.getValue();
                                if (value instanceof String) {
                                    //文本替换
                                    text = text.replace(key, value.toString());
                                } else if (value instanceof Map) {
                                    //图片替换
                                    text = text.replace(key, "");
                                    Map pic = (Map) value;
                                    int width = Integer.parseInt(pic.get("width").toString());
                                    int height = Integer.parseInt(pic.get("height").toString());
                                    int picType = getPictureType(pic.get("type").toString());
                                    //获取图片流，因本人项目中适用流
                                    //InputStream is = (InputStream) pic.get("content");
                                    String byteArray = (String) pic.get("content");
                                    CTInline inline = run.getCTR().addNewDrawing().addNewInline();
                                    /**
                                     * 由于word模板插入图片会导致转pdf时失败报错：Value for parameter 'id' was out of bounds
                                     * 解决办法是在Microsoft Word中另存为后，再转pdf，Linux无法实现这一步骤（不带图没问题，可以直接将填充好的docx转为pdf）
                                     * 原因是，本方法word模板本质为xml，所以填充后生成的word本质仍然为xml，无法实现带图转pdf
                                     * 所以，在填充数据时不填充图片信息，改为转pdf后，再以关键字绝对坐标方式向指定位置插入图片
                                     */
                                    // insertPicture(doc, byteArray, inline, width, height,picType);
                                }
                            }
                        }
                        if (isSetText) {
                            run.setText(text, 0);
                        }
                    }
                }
            }
        }
    }

    // oss地址流式下载
    public static InputStream getObjectStream(String path) {

        OSSClient client = new OSSClient("https://oss-cn-north-2-gov-1.aliyuncs.com", "LTAI4FiGsPymakTEmDUmpWjS", "oHV2tpIHHIGAnR0fCTajexUQWdDEVD");
        OSSObject object = client.getObject(new GetObjectRequest("super-app-01", path));
        InputStream objectContent = object.getObjectContent();
        return objectContent;
    }

    /**
     * 插入图片
     */
    private static void insertPicture(XWPFDocument document, String filePath,
                                      CTInline inline, int width,
                                      int height,int imgType) throws InvalidFormatException, FileNotFoundException {

        //通过流获取图片，因本人项目中，是通过流获取
        document.addPictureData(getObjectStream(filePath),imgType);
        int id = document.getAllPictures().size() - 1;
        final int EMU = 9525;
        width *= EMU;
        height *= EMU;
        String blipId = document.getRelationId(document.getAllPictures().get(id));
        String picXml = getPicXml(blipId, width, height);
        XmlToken xmlToken = null;
        try {
            xmlToken = XmlToken.Factory.parse(picXml);
        } catch (XmlException xe) {
            xe.printStackTrace();
        }
        inline.set(xmlToken);
        inline.setDistT(0);
        inline.setDistB(0);
        inline.setDistL(0);
        inline.setDistR(0);
        CTPositiveSize2D extent = inline.addNewExtent();
        extent.setCx(width);
        extent.setCy(height);
        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
        docPr.setId(id);
        docPr.setName("IMG_" + id);
        docPr.setDescr("IMG_" + id);
    }


    /**
     * 根据图片类型，取得对应的图片类型代码
     *
     * @param picType
     * @return int
     */
    private static int getPictureType(String picType) {

        int res = CustomXWPFDocument.PICTURE_TYPE_PICT;
        if (picType != null) {

            if (picType.equalsIgnoreCase("png")) {

                res = CustomXWPFDocument.PICTURE_TYPE_PNG;
            } else if (picType.equalsIgnoreCase("dib")) {

                res = CustomXWPFDocument.PICTURE_TYPE_DIB;
            } else if (picType.equalsIgnoreCase("emf")) {

                res = CustomXWPFDocument.PICTURE_TYPE_EMF;
            } else if (picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")) {

                res = CustomXWPFDocument.PICTURE_TYPE_JPEG;
            } else if (picType.equalsIgnoreCase("wmf")) {

                res = CustomXWPFDocument.PICTURE_TYPE_WMF;
            }
        }
        return res;
    }


    private static String getPicXml(String blipId, int width, int height) {

        String picXml =
                "" + "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                        "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                        "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                        "         <pic:nvPicPr>" + "            <pic:cNvPr id=\"" + 0 +
                        "\" name=\"Generated\"/>" + "            <pic:cNvPicPr/>" +
                        "         </pic:nvPicPr>" + "         <pic:blipFill>" +
                        "            <a:blip r:embed=\"" + blipId +
                        "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
                        "            <a:stretch>" + "               <a:fillRect/>" +
                        "            </a:stretch>" + "         </pic:blipFill>" +
                        "         <pic:spPr>" + "            <a:xfrm>" +
                        "               <a:off x=\"0\" y=\"0\"/>" +
                        "               <a:ext cx=\"" + width + "\" cy=\"" + height +
                        "\"/>" + "            </a:xfrm>" +
                        "            <a:prstGeom prst=\"rect\">" +
                        "               <a:avLst/>" + "            </a:prstGeom>" +
                        "         </pic:spPr>" + "      </pic:pic>" +
                        "   </a:graphicData>" + "</a:graphic>";
        return picXml;
    }

    public static void main(String[] args) throws Exception {

        Map<String,Object> param = new HashMap<>();
        param.put("1","合规性分析测试");
        param.put("2","总体规划");
        param.put("3","25438834.17");
        param.put("4","0.00");
        param.put("5","执照");
        // param.put("15","执照");
        Map<String,Object> header = new HashMap<>();
        header.put("width", 100);
        header.put("height", 45);
        header.put("type", "jpg");
        header.put("content", "lx/images/fission-image/user/children22.jpg");//图片路径
        param.put("15",header);
        XWPFDocument doc = WordUtil.generateWord(param, "C:\\Users\\87172\\Documents\\培训服务合同（三方）填充模板.docx");
        FileOutputStream fopts = new FileOutputStream("C:\\Users\\87172\\Downloads\\培训服务合同（三方）测试填充.docx");
        doc.write(fopts);
        fopts.close();
        InputStream source = new FileInputStream("C:\\Users\\87172\\Downloads\\培训服务合同（三方）测试填充.docx");
        OutputStream target = new FileOutputStream("C:\\Users\\87172\\Downloads\\培训服务合同（三方）填充PDF.pdf");
        Map<String, String> params = new HashMap<>();

        PdfOptions options = PdfOptions.create();

        WordToPDF.wordConverterToPdf(source, target, options, params);
        // 获取所有图片
        List<String> imglist = new ArrayList<>();
        imglist.add("C:\\Users\\87172\\Pictures\\BFFEC115-FA58-4629-A85C-6AA8AA518651.png");
        PDFInsertImgUtils.addPdfMark("C:\\Users\\87172\\Downloads\\培训服务合同（三方）填充PDF.pdf","C:\\Users\\87172\\Downloads\\图片PDF.pdf",imglist);
    }
}