// Code in Java to convert XLS file to CSV.
// Each xls sheet is written in separate file.
// Reference: http://www.java-tips.org/

import java.io.*;

import jxl.*;

import java.util.*;

class  ConvertCSV
{
  public static void main(String[] args) 
  {
    try
    {
      //File to store data in form of CSV
    	String fiName = "C:\\Users\\Dell\\Desktop\\Projects\\MacroEconomy\\GettingData\\DaneM_nSheets.csv";
    	String encoding = "UTF-16";
    	

      //Excel document to be imported
      String filename = "C:\\Users\\Dell\\Desktop\\Projects\\MacroEconomy\\GettingData\\input.xls";
      WorkbookSettings ws = new WorkbookSettings();
      ws.setLocale(new Locale("en", "EN"));
      Workbook w = Workbook.getWorkbook(new File(filename),ws);

      // Gets the sheets from workbook
      for (int sheet = 0; sheet < w.getNumberOfSheets(); sheet++)
      {
    	 
    	  String nR = Integer.toString(sheet);
    	  String fName = fiName.replaceAll("nSheets",nR);
    	  
      	File f = new File(fName);

        OutputStream os = (OutputStream)new FileOutputStream(f);
        
        OutputStreamWriter osw = new OutputStreamWriter(os, encoding);
        BufferedWriter bw = new BufferedWriter(osw);
 
    	  
        Sheet s = w.getSheet(sheet);

        bw.write(s.getName());
        bw.newLine();

        Cell[] row = null;
        
        // Gets the cells from sheet
        for (int i = 0 ; i < s.getRows() ; i++)
        {
          row = s.getRow(i);

          if (row.length > 0)
          {
            bw.write(row[0].getContents());
            for (int j = 1; j < row.length; j++)
            {
              bw.write(',');
              bw.write(row[j].getContents());
            }
          }
          bw.newLine();
        }
        bw.flush();
        bw.close();

      }
    }
    catch (UnsupportedEncodingException e)
    {
      System.err.println(e.toString());
    }
    catch (IOException e)
    {
      System.err.println(e.toString());
    }
    catch (Exception e)
    {
      System.err.println(e.toString());
    }
  }
}
