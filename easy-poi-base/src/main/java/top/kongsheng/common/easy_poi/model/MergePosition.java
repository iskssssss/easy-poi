package top.kongsheng.common.easy_poi.model;

/**
 * MergePosition
 *
 * @author 孔胜
 * @version 版权 Copyright(c)2024 KONG SHENG
 * @date 2022/8/23 9:22
 */
public class MergePosition {

    private String mergeText;
    private int startY = -1;
    private int endY = -1;
    private int startX = -1;
    private int endX = -1;

    public String getMergeText() {
        return mergeText;
    }

    public void setMergeText(String mergeText) {
        this.mergeText = mergeText;
    }

    public int getStartY() {
        return startY;
    }

    public void setStartY(int startY) {
        this.startY = startY;
    }

    public int getEndY() {
        return endY;
    }

    public void setEndY(int endY) {
        this.endY = endY;
    }

    public int getStartX() {
        return startX;
    }

    public void setStartX(int startX) {
        this.startX = startX;
    }

    public int getEndX() {
        return endX;
    }

    public void setEndX(int endX) {
        this.endX = endX;
    }

    public void setAutoY(int y) {
        if (startY == -1) {
            startY = y;
            return;
        }
        endY = y;
    }

    public void setX(int x) {
        startX = x;
        endX = x;
    }
}
