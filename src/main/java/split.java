import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.Scanner;
public class split {
    private Sheet sheet;
    private XSSFWorkbook wb;
    public void savetofile() throws IOException
    {
        FileOutputStream outFile =new FileOutputStream(new File("update.xlsx"));
        wb.write(outFile);
        outFile.close();
    }
    public void gettowork() throws IOException{
        int columncount=0;
        int rowcount=0;
        String[] sınıflar = {"A","B","C","D","E","F","G","H","I","J","K","L","M","O","P","R","S","T","U","V","Y","Z","X"};
        Scanner in = new Scanner(System.in);
        System.out.print("kaçlı gruplara bölmek istiyorsun: ");
        int dv = in.nextInt();
        if (dv<15 || dv>30)
        {
            System.out.print("HATA! en az 15 kişi, en fazla 30 kişi yazman gerek:");
            dv = in.nextInt();
        }
        FileInputStream inputStream = new FileInputStream(new File("tablo1.xlsx"));
        wb = new XSSFWorkbook(inputStream);
        sheet = wb.getSheetAt(0);
        rowcount = sheet.getLastRowNum();
        //Cell cell = sheet.getRow(1).createCell(6);
        //cell.setCellValue("TEST");

        int i=2;
        int sınıfsırası=0;
        int odasırası=1;
        while (i<rowcount)
        {

             Cell cell = sheet.getRow(i).createCell(6);
             Cell username=sheet.getRow(i).getCell(3);
             Cell surname= sheet.getRow(i).getCell(2);
             if (surname==null)
             {
              surname=sheet.getRow(i).createCell(2);
              //TODO emailden soyadını düzelme
              //TODO olabiliyorsa türkçe düzeltme

             }
             if (username==null)
             {
                 username=sheet.getRow(i).createCell(3);
                 String temp;
                 temp = sheet.getRow(i).getCell(14).getStringCellValue();
                 System.out.println(temp);
                 int index = temp.indexOf('@');
                 temp = temp.substring(0,index);
                 temp = temp.replace(".","");
                 username.setCellValue(temp);
             }
             String val = sınıflar[sınıfsırası]+""+odasırası;
             cell.setCellValue(val);
            if((i%dv)==0)
            {
                odasırası++;
            }
            if ((odasırası % 51)==0)
            {
                sınıfsırası++;
                odasırası=1;
            }
            i++;
            //TODO son satıra sınıf atamıyor
        }
        savetofile();

    }
}
