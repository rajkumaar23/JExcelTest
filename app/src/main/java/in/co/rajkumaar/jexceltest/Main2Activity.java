package in.co.rajkumaar.jexceltest;

import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.widget.Toast;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class Main2Activity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main2);
        Workbook workbook = null;
        File file = new File(Environment.getExternalStorageDirectory() + "/"
                + "AmritaRepo/"+"excelFile.xls");
        try {

            workbook = Workbook.getWorkbook(file);

            Sheet sheet = workbook.getSheet(0);
            Cell cell1 = sheet.getCell(0, 0);
            Toast.makeText(this,cell1.getContents() + ":",Toast.LENGTH_SHORT).show();    // Test Count + :
            Cell cell2 = sheet.getCell(0, 1);
            Toast.makeText(this,cell2.getContents() + ":",Toast.LENGTH_SHORT).show();
            //System.out.println(cell2.getContents());        // 1

            Cell cell3 = sheet.getCell(1, 0);
            Toast.makeText(this,cell3.getContents() + ":",Toast.LENGTH_SHORT).show();
            //System.out.print(cell3.getContents() + ":");    // Result + :
            Cell cell4 = sheet.getCell(1, 1);
            Toast.makeText(this,cell4.getContents() + ":",Toast.LENGTH_SHORT).show();
            //System.out.println(cell4.getContents());        // Passed

            Toast.makeText(this,cell1.getContents() + ":",Toast.LENGTH_SHORT).show();
            //System.out.print(cell1.getContents() + ":");    // Test Count + :
            cell2 = sheet.getCell(0, 2);
            Toast.makeText(this,cell2.getContents() + ":",Toast.LENGTH_SHORT).show();
            //System.out.println(cell2.getContents());        // 2
            Toast.makeText(this,cell3.getContents() + ":",Toast.LENGTH_SHORT).show();
            //System.out.print(cell3.getContents() + ":");    // Result + :
            cell4 = sheet.getCell(1, 2);
            Toast.makeText(this,cell4.getContents() + ":",Toast.LENGTH_SHORT).show();
            //System.out.println(cell4.getContents());        // Passed 2

        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } finally {

            if (workbook != null) {
                workbook.close();
            }

        }

    }

}
