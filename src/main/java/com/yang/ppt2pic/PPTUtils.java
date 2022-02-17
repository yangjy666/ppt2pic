package com.yang.ppt2pic;

import cn.hutool.core.io.IoUtil;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * ppt处理工具
 * @author yangjy
 * @date 2022/02/17
 */
public class PPTUtils {

    /**
     * 将PPT 文件转换成image
     *
     * @param originalPPTFileName //PPT文件路径 如：d:/demo/demo1.ppt
     * @param targetDir //转换后的图片保存路径 如：d:/demo/pptImg
     * @param imageFormat //图片转化的格式字符串 ，如："jpg"、"jpeg"、"bmp" "png" "gif" "tiff"
     * @param times 生成图片放大的倍数,倍数越高，清晰度越高
     * @return 图片名列表
     */
    @SuppressWarnings("resource")
    public static List<String> convertPPTtoImage(
            String originalPPTFileName, String targetDir, String imageFormat, int times) {

        long start = System.currentTimeMillis();
        File dirFile = new File(targetDir);
        if(!dirFile.exists()){
            dirFile.mkdirs();
        }

        int isPpt = checkPPTSuffix(originalPPTFileName);
        if (isPpt < 0) {
            System.out.println("this file isn't ppt or pptx.");
            return null;
        }

        List<String> images = new ArrayList<>();
        FileInputStream input = null;
        //创建文件夹
        createDirIfNotExist(targetDir);

        try {
            input = new FileInputStream(originalPPTFileName);
            if(isPpt == 0) {
                dealPpt(input, targetDir, imageFormat, times, images);
            }else if(isPpt == 1) {
                dealPptx(input, targetDir, imageFormat, times, images);
            }

            IoUtil.close(input);
            System.out.println("completed in " + (System.currentTimeMillis() - start)+ " ms.");
        } catch (IOException e) {
            e.printStackTrace();
            return Collections.emptyList();
        }finally {
            IoUtil.close(input);
        }

        return images;
    }

    private static void dealPpt(FileInputStream input, String targetDir,String imageFormat, int times, List<String> images){

        try {
            HSLFSlideShow ppt = new HSLFSlideShow(input);
            // 获取PPT每页的大小(宽和高度)
            Dimension onePPTPageSize = ppt.getPageSize();
            // 获得PPT文件中的所有的PPT页面(获得每一张幻灯片),并转为一张张的播放片
            List<HSLFSlide> pptPageSlideList = ppt.getSlides();
            // 下面循环的主要功能是实现对PPT文件中的每一张幻灯片进行转换和操作
            for (int i = 0; i < pptPageSlideList.size(); i++) {
                // 这几个循环只要是设置字体为宋体，防止中文乱码，
                List<List<HSLFTextParagraph>> oneTextParagraphs = pptPageSlideList.get(i).getTextParagraphs();

                for (List<HSLFTextParagraph> list : oneTextParagraphs) {
                    for (HSLFTextParagraph hslfTextParagraph : list) {
                        List<HSLFTextRun> HSLFTextRunList = hslfTextParagraph.getTextRuns();
                        for (HSLFTextRun hslfTextRun : HSLFTextRunList) {
                            // 如果PPT在WPS中保存过，则
                            // HSLFTextRunList.get(j).getFontSize();的值为0或者26040，
                            // 因此首先识别当前文本框内的字体尺寸是否为0或者大于26040，则设置默认的字体尺寸。
                            // 设置字体大小
                            Double size = hslfTextRun.getFontSize();
                            if ((size <= 0) || (size >= 26040)) {
                                hslfTextRun.setFontSize(20.0);
                            }

                            // 设置字体样式为宋体
                            hslfTextRun.setFontFamily("宋体");
                        }
                    }
                }

                //创建BufferedImage对象，图像的尺寸为原来的每页的尺寸*倍数times
                BufferedImage oneBufferedImage = new BufferedImage(onePPTPageSize.width * times,
                        onePPTPageSize.height * times, BufferedImage.TYPE_INT_RGB);
                Graphics2D oneGraphics2D = oneBufferedImage.createGraphics();

                // 设置转换后的图片背景色为白色
                oneGraphics2D.setPaint(Color.white);
                oneGraphics2D.scale(times, times);// 将图片放大times倍
                oneGraphics2D.fill(new Rectangle2D.Float(0, 0, onePPTPageSize.width * times, onePPTPageSize.height * times));
                pptPageSlideList.get(i).draw(oneGraphics2D);

                //设置图片的存放路径和图片格式，注意生成的图片路径为绝对路径，最终获得各个图像文件所对应的输出流对象
                //转换后的图片文件保存的指定的目录中
                FileOutputStream out = null;
                try {
                    String imgName = (i + 1) + "." + imageFormat;
                    images.add(imgName);
                    out = new FileOutputStream(targetDir + imgName);
                    javax.imageio.ImageIO.write(oneBufferedImage, "png", out);
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } finally {
                    IoUtil.close(out);
                }

            }

        } catch (IOException ioException) {
            ioException.printStackTrace();
        } finally {
            IoUtil.close(input);
        }

    }

    private static void dealPptx(FileInputStream input, String targetDir,String imageFormat, int times, List<String> images){
        try {
            XMLSlideShow pptx = new XMLSlideShow(input);
            // 获取PPT每页的大小(宽和高度)
            Dimension onePPTPageSize = pptx.getPageSize();
            // 获得PPT文件中的所有的PPT页面(获得每一张幻灯片),并转为一张张的播放片
            List<XSLFSlide> pptPageSlideList = pptx.getSlides();

            // 下面循环的主要功能是实现对PPT文件中的每一张幻灯片进行转换和操作
            for (int i = 0; i < pptPageSlideList.size(); i++) {
                // 这几个循环只要是设置字体为宋体，防止中文乱码，
                for(XSLFShape shape : pptPageSlideList.get(i).getShapes()){
                    if(shape instanceof XSLFTextShape) {
                        XSLFTextShape tsh = (XSLFTextShape)shape;
                        for(XSLFTextParagraph p : tsh){
                            for(XSLFTextRun r : p){
                                // 如果PPT在WPS中保存过，则
                                // HSLFTextRunList.get(j).getFontSize();的值为0或者26040，
                                // 因此首先识别当前文本框内的字体尺寸是否为0或者大于26040，则设置默认的字体尺寸。
                                // 设置字体大小
                                Double size = r.getFontSize();
                                if ((size <= 0) || (size >= 26040)) {
                                    r.setFontSize(20.0);
                                }
                                r.setFontFamily("宋体");
                            }
                        }
                    }
                }

                //创建BufferedImage对象，图像的尺寸为原来的每页的尺寸*倍数times
                BufferedImage oneBufferedImage = new BufferedImage(onePPTPageSize.width * times,
                        onePPTPageSize.height * times, BufferedImage.TYPE_INT_RGB);
                Graphics2D oneGraphics2D = oneBufferedImage.createGraphics();

                // 设置转换后的图片背景色为白色
                oneGraphics2D.setPaint(Color.white);
                oneGraphics2D.scale(times, times);// 将图片放大times倍
                oneGraphics2D.fill(new Rectangle2D.Float(0, 0, onePPTPageSize.width * times, onePPTPageSize.height * times));
                pptPageSlideList.get(i).draw(oneGraphics2D);

                //设置图片的存放路径和图片格式，注意生成的图片路径为绝对路径，最终获得各个图像文件所对应的输出流对象
                //转换后的图片文件保存的指定的目录中
                FileOutputStream out = null;
                try {
                    String imgName = (i + 1) + "." + imageFormat;
                    images.add(imgName);
                    out = new FileOutputStream(targetDir + imgName);
                    javax.imageio.ImageIO.write(oneBufferedImage, "png", out);
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } finally {
                    IoUtil.close(out);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            IoUtil.close(input);
        }
    }

    /**
     * 创建文件如果路径不存在则创建对应的文件夹
     * @param file 文件
     * @return File
     */
    public static File createDirIfNotExist(String file) {

        File fileDir = new File(file);
        if (!fileDir.exists()) {
            fileDir.mkdirs();
        }

        return fileDir;

    }

    public static int checkPPTSuffix(String fileName) {
        int isPpt = -1;
        String suffixName;
        if (fileName != null && fileName.contains(".")) {
            suffixName = fileName.substring(fileName.indexOf("."));
            if (suffixName.equals(".ppt")) {
                isPpt = 0;
            }else if (suffixName.equals(".pptx")) {
                isPpt = 1;
            }
        }

        return isPpt;
    }

    public static void main(String[] args) {
        List<String> result = convertPPTtoImage("C:/Users/yang/Desktop/述职.pptx", "C:/Users/yang/Desktop/pic/", "png", 3);
//        List<String> result = convertPPTtoImage("C:/Users/yang/Desktop/[周越]转正述职.ppt", "C:/Users/yang/Desktop/pic/", "png", 3);

        for(String s : result){
            System.out.println(s);
        }
    }
}
