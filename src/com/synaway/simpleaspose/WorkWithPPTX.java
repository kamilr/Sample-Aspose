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
import com.aspose.slides.SlidesEx;
import com.aspose.slides.Collections.ICollection;



public class WorkWithPPTX {

    public static void main(String[] args) throws Exception{
        String dataDir = "Resources/";
        License license = new License();

        PresentationEx presopen = new PresentationEx(dataDir + "Level_3_VP_Tool_Report_Template.pptx");
        
        System.out.println("slides count is "+Integer.toString(presopen.getSlides().size()));

        processPresentation(presopen);
   
//        presopen.write(dataDir + "toFile.pptx");
//
//        System.out.println("Presentation saved to a file successfully.");
//        ___________________________________
        try
        {
            FileOutputStream outStream = new FileOutputStream(dataDir + "toStream2.pptx");

            presopen.write(outStream);
            outStream.close();

            System.out.println("Presentation saved to a file stream successfully.");
        }
        catch (Exception e)
        {
            System.out.println(e.toString());
        }
//        ________________________________-
   }
    public static void processPresentation(PresentationEx pres)
    {
        int lastSlideNr = pres.getSlides().size();
        
        LayoutSlideEx arg0;

        SlideEx sld = pres.getSlides().get_Item(0);

        int idx = sld.getShapes().addAutoShape(ShapeTypeEx.Rectangle, 150, 75,150 , 50);
        AutoShapeEx ashp = (AutoShapeEx)sld.getShapes().get_Item(idx);

        ashp.addTextFrame("Hello Synaway's member");

        ashp.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getFillFormat().setFillType(FillTypeEx.Solid);
        ashp.getTextFrame().getParagraphs().get_Item(0).getPortions().get_Item(0).getFillFormat().getSolidFillColor().setColor(java.awt.Color.red);

        ashp.getShapeStyle().getLineColor().setColor(java.awt.Color.green);

        ashp.getFillFormat().setFillType(FillTypeEx.NoFill);
    }
}
