package com.example.demoexcelfile;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import android.Manifest;
import android.app.Activity;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.Toast;

/*import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;*/

import com.google.android.gms.auth.UserRecoverableAuthException;
import com.google.android.gms.auth.UserRecoverableNotifiedException;
import com.google.android.gms.auth.api.signin.GoogleSignIn;
import com.google.android.gms.auth.api.signin.GoogleSignInClient;
import com.google.android.gms.auth.api.signin.GoogleSignInOptions;
import com.google.android.gms.common.SignInButton;
import com.google.android.gms.common.api.Scope;
import com.google.android.gms.tasks.OnFailureListener;
import com.google.android.gms.tasks.OnSuccessListener;
import com.google.api.client.extensions.android.http.AndroidHttp;
import com.google.api.client.googleapis.extensions.android.gms.auth.GoogleAccountCredential;
import com.google.api.client.googleapis.extensions.android.gms.auth.UserRecoverableAuthIOException;
import com.google.api.client.http.FileContent;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.Locale;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {

    private Button btn_create;
    private Button btn_uploadFile;

    private WritableCellFormat timesBoldUnderline;
    private WritableCellFormat times;
    private String inputFile;
    ///
    private File excelFile;

    ////
    private static final String TAG = "tejadroid-quickstart";
    private static final int REQUEST_CODE_SIGN_IN = 0;
    private static final int REQUEST_AUTHORIZATION =11;

    private DriveServiceHelper mDriveServiceHelper;

    private GoogleSignInClient mGoogleSignInClient;
    private SignInButton signInButton;
    private Drive driveService;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl");
        btn_create = findViewById(R.id.btn_create);
        btn_uploadFile = findViewById(R.id.btn_uploadFile);
        btn_uploadFile.setOnClickListener(this);
        askWritePermission();
        registerSignInButton();
        buildGGSignInClient();
        btn_create.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                setOutputFile();
                try{

                    createExcelFile();
                }catch (IOException e){
                    e.printStackTrace();
                }catch (WriteException e){
                    e.printStackTrace();
                }
                Toast.makeText(getApplicationContext(),"Complete",Toast.LENGTH_SHORT).show();
            }
        });
    }
    public void registerSignInButton(){
        signInButton = findViewById(R.id.sign_in_button);
        signInButton.setSize(SignInButton.SIZE_STANDARD);
        signInButton.setOnClickListener(this);
    }

    @Override
    public void onClick(View v) {
        switch (v.getId()){
            case R.id.sign_in_button: signIn();break;
            case R.id.btn_uploadFile: uploadFileToGGDrive();break;
        }
    }

    public void signIn(){
        Intent signInIntent = mGoogleSignInClient.getSignInIntent();
        startActivityForResult(signInIntent,REQUEST_CODE_SIGN_IN);
    }

    public void uploadFileToGGDrive(){
        if(mDriveServiceHelper!=null){
            Log.d(TAG,"Creating a file");
            mDriveServiceHelper.createFile(excelFile).addOnSuccessListener(new OnSuccessListener<String>() {
                @Override
                public void onSuccess(String s) {
                    Log.d(TAG,"ID: "+s);
                }
            }).addOnFailureListener(new OnFailureListener() {
                @Override
                public void onFailure(@NonNull Exception e) {

                    Log.e(TAG, "Couldn't create file.",e);
                }
            });
        }else {
            Log.d(TAG,"Fail");
        }
    }

   /*private void readFile(String fileId) {
        if (mDriveServiceHelper != null) {
            Log.d(TAG, "Reading file " + fileId);

            mDriveServiceHelper.readFile(fileId)
                    .addOnSuccessListener(nameAndContent -> {
                        String name = nameAndContent.first;
                        String content = nameAndContent.second;

                        mFileTitleEditText.setText(name);
                        mDocContentEditText.setText(content);

                        setReadWriteMode(fileId);
                    })
                    .addOnFailureListener(exception ->
                            Log.e(TAG, "Couldn't read file.", exception));
        }
    }*/

    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {
        switch (requestCode){
            case REQUEST_CODE_SIGN_IN:
                if(resultCode==Activity.RESULT_OK && data!=null){
                    handleSignInResult(data);
                }
        }
        super.onActivityResult(requestCode,resultCode,data);
    }


    private void handleSignInResult(Intent result) {
        GoogleSignIn.getSignedInAccountFromIntent(result)
                .addOnSuccessListener(googleAccount -> {
                    Log.d(TAG, "Signed in as " + googleAccount.getEmail());

                    // Use the authenticated account to sign in to the Drive service.
                    GoogleAccountCredential credential =
                            GoogleAccountCredential.usingOAuth2(
                                    this, Collections.singleton(DriveScopes.DRIVE_FILE));

                    credential.setSelectedAccount(googleAccount.getAccount());

                    Drive googleDriveService =
                            new Drive.Builder(
                                    AndroidHttp.newCompatibleTransport(),
                                    new GsonFactory(),
                                    credential)
                                    .setApplicationName("Drive API Migration")
                                    .build();
                    // The DriveServiceHelper encapsulates all REST API and SAF functionality.
                    // Its instantiation is required before handling any onClick actions.
                    mDriveServiceHelper = new DriveServiceHelper(googleDriveService);
                })
                .addOnFailureListener(exception -> Log.e(TAG, "Unable to sign in.", exception));
    }

    public void buildGGSignInClient(){
        GoogleSignInOptions gso = new GoogleSignInOptions.Builder(GoogleSignInOptions.DEFAULT_SIGN_IN)
                .requestEmail()
                .requestScopes(new Scope(DriveScopes.DRIVE_FILE))
                .build();
        mGoogleSignInClient = GoogleSignIn.getClient(this,gso);
    }


    public void askWritePermission(){
        if(ActivityCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE)!= PackageManager.PERMISSION_GRANTED){
            ActivityCompat.requestPermissions(this,new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE},28);
        }
    }

    public void setOutputFile() {

        String fileName = "anlele.xls";
        //  String extStorageDirectory= getApplicationContext().getExternalFilesDir(null).getAbsolutePath();
        File folder = new File(getExternalFilesDir(null),"ExcelFolder");
        folder.mkdir();
        excelFile = new File(folder,fileName);
        try{
            excelFile.createNewFile();
            this.inputFile = excelFile.getAbsolutePath();
        }catch (IOException e){
            e.printStackTrace();
        }

    }

    public void upLoadFile() throws IOException {
        com.google.api.services.drive.model.File fileMetaData = new com.google.api.services.drive.model.File();
        fileMetaData.setName("anlele.xls");
        File filePath = new File(excelFile.getAbsolutePath());
        FileContent mediaContent = new FileContent("application/vnd.ms-excel", filePath);
        com.google.api.services.drive.model.File file = driveService.files().create(fileMetaData,mediaContent)
                .setFields("id")
                .execute();
        System.out.println("File ID: " + file.getId());
    }

    public void createExcelFile() throws IOException, WriteException {
        File file = new File(inputFile);
        WorkbookSettings wbSettings = new WorkbookSettings();
        wbSettings.setLocale(new Locale("en","EN"));

        WritableWorkbook workbook = Workbook.createWorkbook(file,wbSettings);
        workbook.createSheet("Report",0);
        WritableSheet excelSheet = workbook.getSheet(0);
        createLabel(excelSheet);
        createContent(excelSheet);

        workbook.write();
        workbook.close();
        //Creating A Sheet
        /*Workbook workbook = new XSSFWorkbook();

        try {

            Sheet sheet = workbook.createSheet("AnLe");
            Cell cell;
            Row row = sheet.createRow(0);
            //Inserting Value in the Sheet
            for (int k = 0; k < 3; k++) {
                cell = row.createCell(k);
                cell.setCellValue(arrData[k]);
                sheet.setColumnWidth(k, 400);
            }
        }catch (Exception e){
            Toast.makeText(this,e.getMessage(),Toast.LENGTH_SHORT).show();
        }



        String fileName = "AHIHI.xlsx";
      //  String extStorageDirectory= getApplicationContext().getExternalFilesDir(null).getAbsolutePath();
        File folder = new File(getExternalFilesDir(null),"ExcelFolder");
        folder.mkdir();
        File file = new File(folder,fileName);
        try{
            file.createNewFile();
        }catch (IOException e){
            e.printStackTrace();
        }

        try{
            FileOutputStream fileOut = new FileOutputStream(file);
            workbook.write(fileOut);
            fileOut.close();
        }catch (FileNotFoundException e){
            e.printStackTrace();
        }catch (IOException e){
            e.printStackTrace();
        }*/

    }

    private void createLabel(WritableSheet sheet)
            throws WriteException {
        // Lets create a times font
        WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);
        // Define the cell format
        times = new WritableCellFormat(times10pt);
        // Lets automatically wrap the cells
        times.setWrap(true);

        // create create a bold font with unterlines
        WritableFont times10ptBoldUnderline = new WritableFont(
                WritableFont.TIMES, 10, WritableFont.BOLD, false,
                UnderlineStyle.SINGLE);
        timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);
        // Lets automatically wrap the cells
        timesBoldUnderline.setWrap(true);

        CellView cv = new CellView();
        cv.setFormat(times);
        cv.setFormat(timesBoldUnderline);
        cv.setAutosize(true);

        // Write a few headers
        addCaption(sheet, 0, 0, "Header 1");
        addCaption(sheet, 1, 0, "This is another header");
    }

    private void createContent(WritableSheet sheet) throws WriteException,
            RowsExceededException {
        // Write a few number
        for (int i = 1; i < 10; i++) {
            // First column
            addNumber(sheet, 0, i, i + 10);
            // Second column
            addNumber(sheet, 1, i, i * i);
        }
        // Lets calculate the sum of it
        StringBuffer buf = new StringBuffer();
        buf.append("SUM(A2:A10)");
        Formula f = new Formula(0, 10, buf.toString());
        sheet.addCell(f);
        buf = new StringBuffer();
        buf.append("SUM(B2:B10)");
        f = new Formula(1, 10, buf.toString());
        sheet.addCell(f);

        // now a bit of text
        for (int i = 12; i < 20; i++) {
            // First column
            addLabel(sheet, 0, i, "Boring text " + i);
            // Second column
            addLabel(sheet, 1, i, "Another text");
        }
    }

    private void addCaption(WritableSheet sheet, int column, int row, String s)
            throws RowsExceededException, WriteException {
        Label label;
        label = new Label(column, row, s, timesBoldUnderline);
        sheet.addCell(label);
    }

    private void addNumber(WritableSheet sheet, int column, int row,
                           Integer integer) throws WriteException, RowsExceededException {
        Number number;
        number = new Number(column, row, integer, times);
        sheet.addCell(number);
    }

    private void addLabel(WritableSheet sheet, int column, int row, String s)
            throws WriteException, RowsExceededException {
        Label label;
        label = new Label(column, row, s, times);
        sheet.addCell(label);
    }

    //////// Upload File to GG Drive



}
