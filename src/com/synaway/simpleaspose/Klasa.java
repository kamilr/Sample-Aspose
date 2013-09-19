package com.synaway.simpleaspose;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import com.aspose.slides.AutoShapeEx;
import com.aspose.slides.FillTypeEx;
import com.aspose.slides.LayoutSlideEx;
import com.aspose.slides.License;
import com.aspose.slides.Presentation;
import com.aspose.slides.PresentationEx;
import com.aspose.slides.Shape;
import com.aspose.slides.ShapeTypeEx;
import com.aspose.slides.Slide;
import com.aspose.slides.SlideEx;
import com.aspose.slides.Slides;
import com.aspose.slides.SlidesEx;
import com.aspose.slides.Collections.ICollection;



public class Klasa {

    public static void main(String[] args) throws Exception{
        String dataDir = "Resources/";
        Presentation pres = new Presentation();
//        License license = new License();
        
        Slide slide = pres.addEmptySlide();
        com.aspose.slides.Rectangle 
        rect = slide.getShapes().addRectangle(2400, 1800, 1000, 500);
//        rect.getLineFormat().setShowLines(false);
        rect.addTextFrame("Hello World");
//        pres.getSlides().removeAt(0);
        pres.write(dataDir + "hello.ppt");
        
        System.out.println("wow, hello.ppt was created");

        Presentation presopen = new Presentation(dataDir + "hello.ppt");
        
        //Printing the total number of slides present in the presentation
        System.out.println("slides count is "+Integer.toString(presopen.getSlides().size()));
   
        PresentationEx pres1 = new PresentationEx();
        pres1 = presopen;
        //....do some work here.....
        processPresentation(pres1);
//
//        //Save your presentation to a file
//        pres1.write(dataDir + "toFile.pptx");
//
//        System.out.println("Presentation saved to a file successfully.");
//        ___________________________________
        try
        {
            FileOutputStream outStream = new FileOutputStream(dataDir + "toStream2.ppt");

            //Saving the presentation to the output stream of Http Response
            pres1.write(outStream);
            outStream.close();

            System.out.println("Presentation saved to a file stream successfully.");
        }
        catch (Exception e)
        {
            System.out.println(e.toString());
        }
//        ________________________________-
   }
    public static void processPresentation(Presentation pres)
    {
        int lastSlideNr = pres.getSlides().size();
        
        Slides slds = pres.getSlides();
       
        ((Presentation) slds).addEmptySlide();
        //Get the first slide
        SlideEx sld = pres.getSlides().get_Item(lastSlideNr+1);

        //Add an AutoShape of Rectangle type
        int idx = sld.getShapes().addAutoShape(ShapeTypeEx.Rectangle, 150, 75,150 , 50);
        AutoShapeEx ashp = (AutoShapeEx)sld.getShapes().get_Item(idx);

        //Add TextFrame to the Rectangle
        ashp.addTextFrame("Hello World");

        //Change the text color to Black (which is White by default)
        ashp.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getFillFormat().setFillType(FillTypeEx.Solid);
        ashp.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getFillFormat().getSolidFillColor().setColor(java.awt.Color.black);

        //Change the line color of the rectangle to White
        ashp.getShapeStyle().getLineColor().setColor(java.awt.Color.white);

        //Remove any fill formatting in the shape
        ashp.getFillFormat().setFillType(FillTypeEx.NoFill);
    }
}
