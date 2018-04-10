/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package label;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.krysalis.barcode4j.impl.code128.Code128Bean;
import org.krysalis.barcode4j.output.bitmap.BitmapCanvasProvider;
import org.krysalis.barcode4j.tools.UnitConv;

/**
 *
 * @author Vince
 */
public class Label {
    
    static FileInputStream fis = null;


    public static void main(String[] args) {
        // TODO code application logic here
        try{
            
            Code128Bean label = new Code128Bean();
            
            final int dpi = 160;

          //Configure the barcode generator
          label.setModuleWidth(UnitConv.in2mm(2.8f / dpi));
          label.setBarHeight(5);
          label.doQuietZone(false);
          label.setFontSize(1.5);

          //Open output file
          File outputFile = new File("Barcode.jpg");

          FileOutputStream out = new FileOutputStream(outputFile);

          BitmapCanvasProvider canvas = new BitmapCanvasProvider(
              out, "image/x-png", dpi, BufferedImage.TYPE_BYTE_BINARY, false, 0);

          //Generate the barcode
          label.generateBarcode(canvas, "2273");

          //Signal end of generation
          canvas.finish();
   
          System.out.println("Bar Code is generated successfully…");
          fis = new FileInputStream("Barcode.jpg");

            XWPFDocument doc = new XWPFDocument(OPCPackage.open("tableLayout.docx"));

            for (XWPFTable tbl : doc.getTables()) {
               for (XWPFTableRow row : tbl.getRows()) {
                  
                  for (XWPFTableCell cell : row.getTableCells()) {
                        if(cell.getColor()==null){
                            cell.removeParagraph(0);
                        
                            XWPFParagraph p = cell.addParagraph();
                            p.setAlignment(ParagraphAlignment.CENTER);

                            XWPFRun run = p.createRun();

                            fis = new FileInputStream("Barcode.jpg");

                            run.addPicture(fis, XWPFDocument.PICTURE_TYPE_JPEG,null, Units.toEMU(128), Units.toEMU(28));
                        }
                          
                  }
               }
            }
        
            doc.write(new FileOutputStream(new File("poi.docx")));
        }catch(Exception e){
            System.out.println(e.toString());
        }
        
    }
    
}
