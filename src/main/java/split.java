import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.Scanner;
public class split {
    private static String filename;
    private XSSFWorkbook wb;
    public String capitalize(String str) {
        if(str == null || str.isEmpty()) {
            return str;
        }

        return str.substring(0, 1).toUpperCase() + str.substring(1);
    }
    //TODO autodetect excel files example: .xlsx add to list
    public void savetofile() throws IOException
    {
        File theDir = new File("output");
        if (!theDir.exists()){
            theDir.mkdirs();
        }
        String path="";
        FileOutputStream outFile = new FileOutputStream("output/"+filename+"_fixed.xlsx");
        wb.write(outFile);
        outFile.close();
    }
    public void gettowork() throws IOException{
        Scanner in = new Scanner(System.in);
        System.out.print("dosya ismi?:");
        filename = in.nextLine();
        FileInputStream inputStream = new FileInputStream(filename+".xlsx");
        wb = new XSSFWorkbook(inputStream);
        Sheet sheet = wb.getSheetAt(0);
        int rowcount = sheet.getLastRowNum();
        int i=1;
        while (i<=rowcount)
        {

             Cell name= sheet.getRow(i).getCell(1);
             Cell username= sheet.getRow(i).getCell(3);
             Cell surname= sheet.getRow(i).getCell(2);
             Cell pass= sheet.getRow(i).getCell(4);
             Cell room=sheet.getRow(i).getCell(6);
             Cell emailadress = sheet.getRow(i).getCell(14);
             String tr = room.getStringCellValue();
             int rlength = tr.length();
             String emailtemp = emailadress.getStringCellValue();



            /* if (surname==null && name==null) //name & surname from mail adress
             {
                 name=sheet.getRow(i).createCell(1);
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
             */
            if ((room != null || room.getStringCellValue() != "") && rlength<3 )
            {
                String troom = room.getStringCellValue();
                char c1,c2;
                c1 = troom.charAt(0);
                c2 = troom.charAt(1);
                troom = c1+"-"+c2;
                troom.trim();
                room.setCellValue(troom);
            }
             if (username == null || username.getStringCellValue() == "")
             {
                 if (emailadress==null)
                 {
                     System.out.println("Email adresi boş hücre:0"+i);
                     break;
                 }
                 else
                     {
                     username = sheet.getRow(i).createCell(3);
                     int index = emailtemp.indexOf('@');
                     emailtemp = emailtemp.substring(0, index);
                     emailtemp = emailtemp.replace(".", "");
                     username.setCellValue(emailtemp);
                 }
             }
             if (surname == null || surname.getStringCellValue() == "")
             {
              surname = sheet.getRow(i).createCell(2);
              name = sheet.getRow(i).getCell(1);
              String surtemp = name.getStringCellValue();
              String[] surarry = surtemp.split(" ");
              String fname = "";
              surtemp = surarry[surarry.length-1];
              int b = surarry.length-1;
              for (int z=0;z<b;z++)
                 {
                     fname += " "+ surarry[z];
                 }
              surtemp=surtemp.trim();
              surname.setCellValue(surtemp);
              fname=fname.trim();
              name.setCellValue(fname);
             }
             //final trim
            name= sheet.getRow(i).getCell(1);
            username= sheet.getRow(i).getCell(3);
            surname= sheet.getRow(i).getCell(2);
            pass= sheet.getRow(i).getCell(4);
            room=sheet.getRow(i).getCell(6);
            emailadress = sheet.getRow(i).getCell(14);
            String fnametrim = "";
            String fsurnametrim = "";
            String fusernametrim = "";
            String froomtrim = "";
            String femailtrim = "";


             fnametrim = name.getStringCellValue();
             fsurnametrim = surname.getStringCellValue();
             fusernametrim = username.getStringCellValue();
             froomtrim = room.getStringCellValue();
             femailtrim = emailadress.getStringCellValue();
             fnametrim.trim();
             fsurnametrim.trim();
             fusernametrim.trim();
             froomtrim.trim();
             femailtrim.trim();
             name.setCellValue(fnametrim);
             surname.setCellValue(fsurnametrim);
             username.setCellValue(fusernametrim);
             room.setCellValue(froomtrim);
             emailadress.setCellValue(femailtrim);
             pass= sheet.getRow(i).createCell(4);
             pass.setCellValue(123456);


            i++;
        }
        savetofile();

    }
}
