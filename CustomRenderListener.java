package com.longshine.superapp.nucleicresult.util;

import com.itextpdf.awt.geom.Rectangle2D;
import com.itextpdf.text.pdf.parser.ImageRenderInfo;
import com.itextpdf.text.pdf.parser.RenderListener;
import com.itextpdf.text.pdf.parser.TextRenderInfo;

/**
 * @Author Chen ZhiZhe
 * @Create 2021.08.17 14:31
 * @Description
 */
public class CustomRenderListener implements RenderListener {

    private float[] pcoordinate = null;

    private String keyWord;

    private int page;

    public int getPage() {
        return page;
    }

    public void setPage(int page) {
        this.page = page;
    }

    public float[] getPcoordinate(){
        return pcoordinate;
    }

    public String getKeyWord() {
        return keyWord;
    }

    public void setKeyWord(String keyWord) {
        this.keyWord = keyWord;
    }

    @Override
    public void beginTextBlock() {}

    @Override
    public void endTextBlock() {}

    @Override
    public void renderImage(ImageRenderInfo arg0) {}

    @Override
    public void renderText(TextRenderInfo textRenderInfo) {
        String text = textRenderInfo.getText();
        if (null != text && text.contains(keyWord)) {
            Rectangle2D.Float boundingRectange = textRenderInfo.getBaseline().getBoundingRectange();
            pcoordinate = new float[3];
            pcoordinate[0] = boundingRectange.x;
            pcoordinate[1] = boundingRectange.y;
            pcoordinate[2] = page;
        }
    }

}