package in.co.rajkumaar.jexceltest;

import android.Manifest;
import android.content.pm.PackageManager;
import android.os.Environment;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;

import java.io.File;
import jxl.Workbook;
import jxl.write.*;
import jxl.write.Number;
import java.io.IOException;

public class MainActivity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        if (ContextCompat.checkSelfPermission(MainActivity.this,
                Manifest.permission.WRITE_EXTERNAL_STORAGE)
                != PackageManager.PERMISSION_GRANTED) {

            ActivityCompat.requestPermissions(MainActivity.this,
                    new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE},
                    1);

        }

        File file = new File(Environment.getExternalStorageDirectory() + "/"
                + "AmritaRepo/"+"excelFile.xls");

            //1. Create an Excel file
            WritableWorkbook myFirstWbook = null;
            try {


                myFirstWbook = Workbook.createWorkbook(file);

                // create an Excel sheet
                WritableSheet excelSheet = myFirstWbook.createSheet("Sheet 1", 0);

                // add something into the Excel sheet
                Label label = new Label(0, 0, "Test Count");
                excelSheet.addCell(label);

                Number number = new Number(0, 1, 1);
                excelSheet.addCell(number);

                label = new Label(1, 0, "Result");
                excelSheet.addCell(label);

                label = new Label(1, 1, "Value A");
                excelSheet.addCell(label);

                number = new Number(0, 2, 2);
                excelSheet.addCell(number);

                label = new Label(1, 2, "Value B");
                excelSheet.addCell(label);

                myFirstWbook.write();

                Log.e("Written xls ","Successs");


            } catch (IOException e) {
                e.printStackTrace();
            } catch (WriteException e) {
                e.printStackTrace();
            } finally {

                if (myFirstWbook != null) {
                    try {
                        myFirstWbook.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    } catch (WriteException e) {
                        e.printStackTrace();
                    }
                }


            }

        }

    }

