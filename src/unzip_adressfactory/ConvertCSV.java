package unzip_adressfactory;

import java.io.*;
import jxl.*;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

class  ConvertCSV
{
  public static void main(String[] args) throws Exception
  {
//      args=new String[2];
//      args[1]="C:\\test_in";
//      args[0]="C:\\test_in\\09.11.2011.xls";
    
      //File to store data in form of CSV
      File f = new File(args[1]);

      OutputStream os = (OutputStream)new FileOutputStream(f);
      String encoding = "ISO-8859-1";
      OutputStreamWriter osw = new OutputStreamWriter(os, encoding);
      BufferedWriter bw = new BufferedWriter(osw);

      //Excel document to be imported
      String filename = args[0];
      WorkbookSettings ws = new WorkbookSettings();
      ws.setEncoding("cp1252");
      ws.setLocale(new Locale("en", "EN"));
      Workbook w = Workbook.getWorkbook(new File(filename),ws);

      // Gets the sheets from workbook
      for (int sheet = 0; sheet < w.getNumberOfSheets(); sheet++)
      {
        Sheet s = w.getSheet(sheet);

       // bw.write(s.getName());
        //bw.newLine();

        Cell[] row = null;
        String acrow="";

        // Gets the cells from sheet
        for (int i = 0 ; i < s.getRows() ; i++)
        {
          row = s.getRow(i);

          if (row.length > 0)
          {
            bw.write(row[0].getContents());
            for (int j = 1; j < row.length; j++)
            {
                acrow=row[j].getContents();
                if(acrow.contains("\n"))
                    acrow=acrow.replaceAll("\n", "");

                if (!( acrow.contains("\n")))
                    
                {
                 bw.write(';');
                 bw.write(acrow);
                }
            }
          }
          bw.newLine();
        }
      }
      bw.flush();
      bw.close();


    
    
  }
}