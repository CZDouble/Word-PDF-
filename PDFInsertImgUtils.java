package com.longshine.superapp.nucleicresult.util;

import com.itextpdf.text.Document;
import com.itextpdf.text.Image;
import com.itextpdf.text.pdf.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.*;

/**
 * @Author Chen ZhiZhe
 * @Create 2021.08.17 13:10
 * @Description pdf插入图片工具类
 */
public class PDFInsertImgUtils {

    /**
     * 该方法下的图片坐标、图片宽高，都需要根据对应pdf内容自行调整，
     * 原因是，图片数据是生成pdf后自行填充，并非模板自动填充，
     * 所以如果图片宽高不固定，不同的图片在页面的位置就会飘忽不定，根据绝对坐标贴的图，必须固定绝对宽高值，也不能按比例缩放
     * @param InPdfFile
     * @param outPdfFile
     * @param imgList
     * @throws Exception
     */
    public static void addPdfMark(String InPdfFile, String outPdfFile, List<String> imgList) throws Exception {
        try {
            PdfReader reader = new PdfReader(InPdfFile, "PDF".getBytes());
            PdfStamper stamp = new PdfStamper(reader, new FileOutputStream(outPdfFile));

            // 关键字坐标以设置图片位置
            float[] keyWordsByPath = PdfHelper.getKeyWordsByPath(InPdfFile, "营业执照");
            Image img = Image.getInstance(imgList.get(0));// 插入图片
            // 设置图片位置。
            img.setAbsolutePosition(310, keyWordsByPath[1] - 23);
            Image img2 = Image.getInstance(imgList.get(0));// 插入图片
            // 设置图片2位置。
            img2.setAbsolutePosition(410, keyWordsByPath[1] - 23);
            // 第一页 如果需要每一页同样位置都加图片，则循环
            PdfContentByte zhizhao1 = stamp.getOverContent(1);
            img.scaleAbsolute(60,35);
            zhizhao1.addImage(img);
            PdfContentByte zhizhao2 = stamp.getOverContent(1);
            // 设置图片绝对宽高
            img2.scaleAbsolute(60,35);
            zhizhao2.addImage(img2);

            stamp.close();// 关闭
            File tempfile = new File(InPdfFile);

            if (tempfile.exists()) {
                tempfile.delete();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static Map<String, Object> getPdfMsg(String filePath) {
        Map<String, Object> map = new LinkedHashMap<String, Object>();
        try {
            // 获取PDF共有几页
            PdfReader pdfReader = new PdfReader(new FileInputStream(filePath));
            int pages = pdfReader.getNumberOfPages();
            // System.err.println(pages);
            map.put("pageSize", pages);

            // 获取PDF 的宽高
            PdfReader pdfreader = new PdfReader(filePath);
            Document document = new Document(pdfreader.getPageSize(pages));
            float widths = document.getPageSize().getWidth();
            // 获取页面高度
            float heights = document.getPageSize().getHeight();
            // System.out.println("widths = " + widths + ", heights = " + heights);
            map.put("width", widths);
            map.put("height", heights);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return map;
    }

    public static Map<String, Object> getImgMsg(String imgPath) {
        Map<String, Object> map = new LinkedHashMap<String, Object>();
        try {
            File picture = new File(imgPath);
            BufferedImage sourceImg = ImageIO.read(new FileInputStream(picture));
            // System.out.println("=源图宽度===>"+sourceImg.getWidth()); // 源图宽度
            // System.out.println("=源图高度===>"+sourceImg.getHeight()); // 源图高度
            map.put("width", sourceImg.getWidth());
            map.put("height", sourceImg.getHeight());
        } catch (Exception e) {
            e.printStackTrace();
        }
        return map;
    }

    public static void main(String[] args) {
        try {
            // 获取所有图片
            List<String> list = new ArrayList<>();
            list.add("C:\\Users\\87172\\Pictures\\BFFEC115-FA58-4629-A85C-6AA8AA518651.png");

            addPdfMark("C:\\Users\\87172\\Downloads\\培训服务合同（三方）填充PDF.pdf", "C:\\Users\\87172\\Downloads\\图片PDF.pdf", list);
        } catch (Exception e) {
            System.out.println("失败");
            e.printStackTrace();
        }
        System.out.println("成功");
    }

}
