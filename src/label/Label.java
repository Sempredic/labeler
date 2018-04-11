/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package label;

import java.awt.Graphics;
import java.awt.image.BufferedImage;
import java.awt.print.PageFormat;
import java.awt.print.Printable;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import javax.print.Doc;
import javax.print.DocFlavor;
import javax.print.DocPrintJob;
import javax.print.PrintService;
import javax.print.SimpleDoc;
import javax.print.attribute.HashDocAttributeSet;
import javax.print.attribute.HashPrintRequestAttributeSet;
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
public class Label implements Printable  {
    
    static FileInputStream fis = null;


    public static void main(String[] args){
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

                            run.addPicture(fis, XWPFDocument.PICTURE_TYPE_JPEG,null, Units.toEMU(105), Units.toEMU(25));
                            run.setText("/18");
                            run.setFontSize(8);
                        }
                          
                  }
               }
            }
        
            doc.write(new FileOutputStream(new File("poi.docx")));
                    
            PrinterJob pj = PrinterJob.getPrinterJob();
        
            if (pj.printDialog()) {
                try {
                    
                    PrintService printer = pj.getPrintService();
                    DocFlavor docType = DocFlavor.INPUT_STREAM.AUTOSENSE;
                    HashDocAttributeSet hashDocAttributeSet = new HashDocAttributeSet();
                    FileInputStream fis = new FileInputStream("poi.doc");
                    Doc wordDoc = new SimpleDoc(fis,docType,hashDocAttributeSet);
                    DocPrintJob printJob = printer.createPrintJob();
                    
                    printJob.print(wordDoc,null);
                    pj.print();
                    fis.close();
              
                    
                }
                catch (PrinterException exc) {
                    System.out.println(exc);
                 }
             }

            
        }catch(Exception e){
            System.out.println(e.toString());
        }
        
    }

    @Override
    public int print(Graphics grphcs, PageFormat pf, int i) throws PrinterException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
    
}
