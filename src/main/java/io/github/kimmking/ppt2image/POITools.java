package io.github.kimmking.ppt2image;

import java.awt.Dimension;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.util.List;
import java.util.LinkedList;

import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 * @Author: kimmking(kimmking@163.com)
 * Convert a PPT or PPTX file to Images by per slide.
 * Created on 2018/02/27.
 */
public class POITools {

    public static void main(String[] args) {
        File file = new File("D:\\git\\PPT2Image\\1.pptx");
        List<String> images = convertPPTtoImage(file,"D:\\git\\PPT2Image\\images\\pptx");

        for ( String s : images){
            System.out.println(s);
        }
    }

    /**
     * @Date: Created in 2018/02/27.
     * @param file the full path of PPT or PPTX file
     * @param dir the full path of directory for saving images
     * @return the list of  image file full names
     */
    public static List<String> convertPPTtoImage(File file, String dir) {

        long start = System.currentTimeMillis();

        File dirFile = new File(dir);
        if(!dirFile.exists()){
            dirFile.mkdirs();
        }

        List<String> images = new LinkedList<String>();
        int isppt = checkFile(file);
        if (isppt < 0) {
            System.out.println("this file isn't ppt or pptx.");
            return null;
        }
        try {
            FileInputStream is = new FileInputStream(file);
            Dimension pgsize = null;
            Object[] slide = null;

            if(isppt == 0) {
                SlideShow ppt = new SlideShow(is);
                is.close();
                pgsize = ppt.getPageSize();
                slide = ppt.getSlides();
            }else if(isppt == 1) {
                XMLSlideShow xppt = new XMLSlideShow(is);
                is.close();
                pgsize = xppt.getPageSize();
                slide = xppt.getSlides();
            }

            for (int i = 0; i < slide.length; i++) {
                System.out.println("Processing Page " + (i+1) + "... ");
                BufferedImage img = new BufferedImage(pgsize.width,
                        pgsize.height, BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();
                graphics.setPaint(Color.white);
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width,
                        pgsize.height));

                if(isppt==0) {
                    Slide ss =  (Slide)slide[i];
                    ss.draw(graphics);
                }else{
                    XSLFSlide xss =  (XSLFSlide)slide[i];
                    xss.draw(graphics);
                }

                String filename = (i + 1) + ".jpg";
                File imgFile = new File(dir ,filename);
                images.add(imgFile.getAbsolutePath());
                FileOutputStream out = new FileOutputStream(imgFile);
                javax.imageio.ImageIO.write(img, "jpg", out);
                out.close();
            }
            System.out.println("completed in " + (System.currentTimeMillis() - start)+ " ms.");
        } catch (FileNotFoundException e) {
            System.out.println(e);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return images;
    }

    public static int checkFile(File file) {
        int isppt = -1;
        String filename = file.getName();
        String suffixname = null;
        if (filename != null && filename.indexOf(".") != -1) {
            suffixname = filename.substring(filename.indexOf("."));
            if (suffixname.equals(".ppt")) {
                isppt = 0;
            }else if (suffixname.equals(".pptx")) {
                isppt = 1;
            }
        }
        return isppt;
    }
}