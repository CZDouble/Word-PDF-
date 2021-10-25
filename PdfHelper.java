package com.longshine.superapp.nucleicresult.util;

import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.parser.PdfReaderContentParser;

import java.io.IOException;

/**
 * @Author Chen ZhiZhe
 * @Create 2021.08.17 14:29
 * @Description 根据pdf关键字获取相应坐标
 */
public class PdfHelper {

    /**
     * @Author AlphaJunS
     * @Date 18:24 2020/3/7
     * @Description 用于供外部类调用获取关键字所在PDF文件坐标
     * @param filepath
     * @param keyWords
     * @return float[]
     */
    public static float[] getKeyWordsByPath(String filepath, String keyWords) {
        float[] coordinate = null;
        try{
            PdfReader pdfReader = new PdfReader(filepath);
            coordinate = getKeyWords(pdfReader, keyWords);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return coordinate;
    }

    /**
     * @Author AlphaJunS
     * @Date 18:26 2020/3/7
     * @Description 获取关键字所在PDF坐标
     * @param pdfReader
     * @param keyWords
     * @return float[]
     */
    private static float[] getKeyWords(PdfReader pdfReader, String keyWords) {
        float[] coordinate = null;
        int page = 0;
        try{
            int pageNum = pdfReader.getNumberOfPages();
            PdfReaderContentParser pdfReaderContentParser = new PdfReaderContentParser(pdfReader);
            CustomRenderListener renderListener = new CustomRenderListener();
            renderListener.setKeyWord(keyWords);
            for (page = 1; page <= pageNum; page++) {
                renderListener.setPage(page);
                pdfReaderContentParser.processContent(page, renderListener);
                coordinate = renderListener.getPcoordinate();
                if (coordinate != null) break;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return coordinate;
    }

}
