import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.Scanner;
public class split {
    private Sheet sheet;
    private XSSFWorkbook wb;
    public String capitalize(String str) {
        if(str == null || str.isEmpty()) {
            return str;
        }

        return str.substring(0, 1).toUpperCase() + str.substring(1);
    }
    public void savetofile() throws IOException
    {
        FileOutputStream outFile = new FileOutputStream("update.xlsx");
        wb.write(outFile);
        outFile.close();
    }
    public void gettowork() throws IOException{
        //int columncount=0;
        int maxroom=0,dv=0;
        String[] classarr = {"A","B","C","D","E","F","G","H","I","J","K","L","M","O","P","R","S","T","U","V","Y","Z","X"};
        Scanner in = new Scanner(System.in);
        do
        {
            System.out.print("grupları hangi sayıya bölmek istiyorsun? (en az 10 en fazla 30): ");
            dv = in.nextInt();

        } while (dv<10 || dv>30);
        do
        {
            System.out.print("bloklardaki max oda sayısı kaç olucak? (en fazla 999): ");
            maxroom = in.nextInt();
            maxroom++;

        } while (maxroom==0 || maxroom>1000 || maxroom<30);
        FileInputStream inputStream = new FileInputStream("tablo1.xlsx");
        wb = new XSSFWorkbook(inputStream);
        sheet = wb.getSheetAt(0);
        int rowcount = sheet.getLastRowNum();
        //Cell cell = sheet.getRow(1).createCell(6);
        //cell.setCellValue("TEST");

        int i=1;
        int classnumber=0;
        int roomnumber=1;
        while (i<rowcount)
        {

             Cell cell = sheet.getRow(i).createCell(6);
             Cell username=sheet.getRow(i).getCell(3);
             Cell surname= sheet.getRow(i).getCell(2);
             if (surname==null)
             {
              surname=sheet.getRow(i).createCell(2);
              String temp = sheet.getRow(i).getCell(14).getStringCellValue();
                 int index = temp.indexOf('@');
                 temp = temp.substring(0,index);
                 temp = temp.split("\\.")[1];
                 temp = temp.replaceAll("[0-9]", "");
                 temp = capitalize(temp);
                 surname.setCellValue(temp);
              //TODO olabiliyorsa türkçe düzeltme

             }
             if (username==null)
             {
                 username=sheet.getRow(i).createCell(3);
                 String temp;
                 temp = sheet.getRow(i).getCell(14).getStringCellValue();
                 //System.out.println(temp);
                 int index = temp.indexOf('@');
                 temp = temp.substring(0,index);
                 temp = temp.replace(".","");
                 username.setCellValue(temp);
             }
             String val = classarr[classnumber]+""+roomnumber;
             cell.setCellValue(val);
            if((i%dv)==0)
            {
                roomnumber++;
            }
            if ((roomnumber % maxroom)==0)
            {
                classnumber++;
                roomnumber=1;
            }
            i++;
            //TODO son satıra sınıf atamıyor
        }
        savetofile();

    }
}
