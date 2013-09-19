package com.synaway.simpleaspose;

import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;

import com.aspose.slides.FillType;
import com.aspose.slides.GradientColorType;
import com.aspose.slides.GradientStyle;
import com.aspose.slides.Presentation;
import com.aspose.slides.Shape;
import com.aspose.slides.Slide;

public class WorkWithPPT {

    public static void main(String[] args) throws Exception {
        String dataDir = "Resources/";
        Presentation pres = new Presentation();
//        slide 1
        Slide slide = pres.addEmptySlide();
        com.aspose.slides.Rectangle 
        rect = slide.getShapes().addRectangle(2400, 1800, 1000, 500);
        rect.getLineFormat().setShowLines(false);
        rect.addTextFrame("Hello World").fitTextToShape();
        pres.getSlides().removeAt(0);
//        slide 2
        pres.addEmptySlide();
        slide = pres.getSlideByPosition(2);
        Shape shape = slide.getShapes().addEllipse(3000, 1200, 1000, 2000);
        shape.getFillFormat().setType(FillType.Gradient);

        shape.getFillFormat().setGradientColorType(GradientColorType.TwoColors);

        shape.getFillFormat().setBackColor(Color.LIGHT_GRAY);
        shape.getFillFormat().setForeColor(Color.ORANGE);

        shape.getFillFormat().setGradientFillAngle((byte)90);

        shape.setRotation((byte)45);
        shape.getLineFormat().setForeColor(Color.red);
//        slide 3
        pres.addEmptySlide();
        slide = pres.getSlideByPosition(3);
        Shape shape2 = slide.getShapes().addPictureFrame(1400, 1100, 3000, 2000, 1200);
        shape2.getFillFormat().setType(FillType.Picture);
        
        slide.setFollowMasterBackground(false);
        slide.getBackground().getFillFormat().setType(FillType.Gradient);
        slide.getBackground().getFillFormat().setGradientColorType(GradientColorType.OneColor);
        slide.getBackground().getFillFormat().setGradientStyle(GradientStyle.FromCenter);
        slide.getBackground().getFillFormat().setForeColor(java.awt.Color.green);
        slide.getBackground().getFillFormat().setGradientDegree((byte)30);
        
        BufferedImage image=null;

        try {
            image = ImageIO.read(new File(dataDir + "Jesien-Autumn-Krajobraz-1.jpg"));
        } catch (IOException e) {
                throw new Exception("Error while opening the image file.");
        }
        com.aspose.slides.Picture
        pic = new com.aspose.slides.Picture(pres,image);
        int picId = pres.getPictures().add(pic);
        shape2.getFillFormat().setPictureId(picId);
        
        
        pres.write(dataDir + "hello.ppt");
        System.out.println("wow, hello.ppt was created");
    }

}
