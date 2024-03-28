package com.motcs.utils.poi.entity;

import lombok.Builder;
import lombok.Data;

import java.io.Serializable;

/**
 * 图片转换类
 *
 * @author <a href="https://github.com/motcs">motcs</a>
 * @since 2024-03-28 星期四
 */
@Data
@Builder
public class ImageUrl implements Serializable {

    /**
     * 图片地址
     */
    private String url;

    /**
     * 图片宽度默认16.2
     */
    private double width;

    /**
     * 图片高度默认10.01
     */
    private double height;

    /**
     * 下载图片错误时的提示语
     */
    private String errMsgD;

    /**
     * 解析图片错误时的提示语
     */
    private String errMsgA;

    /**
     * 解析错误时是否显示图片url
     */
    private Boolean isErrUrl;

    public double getWidth() {
        return width > 0 ? width : 16.2;
    }

    public double getHeight() {
        return height > 0 ? height : 10.01;
    }

    public String getErrMsgD() {
        errMsgD = errMsgD == null ? "解析图片失败!" : errMsgD;
        return errMsgD + (this.getErrUrl() ? this.url : "");
    }

    public String getErrMsgA() {
        errMsgA = errMsgA == null ? "解析图片失败!" : errMsgA;
        return errMsgA + (this.getErrUrl() ? this.url : "");
    }

    public Boolean getErrUrl() {
        return isErrUrl == null || isErrUrl;
    }

}
