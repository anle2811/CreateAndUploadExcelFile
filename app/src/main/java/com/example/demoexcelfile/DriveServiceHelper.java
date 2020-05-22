package com.example.demoexcelfile;

//import com.google.android.gms.drive.Drive;

import android.util.Pair;

import com.google.api.client.http.FileContent;
import com.google.api.services.drive.Drive;

import org.mortbay.thread.Timeout;
import com.google.android.gms.tasks.Task;
import com.google.android.gms.tasks.Tasks;
import com.google.api.services.drive.model.File;


import java.io.BufferedReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Collections;
import java.util.concurrent.Executor;
import java.util.concurrent.Executors;

public class DriveServiceHelper {
    private final Executor mExecutor = Executors.newSingleThreadExecutor();
    private final Drive mDriveService;


    public DriveServiceHelper(Drive driveService) {
        mDriveService = driveService;
    }

    public Task<String> createFile(java.io.File fileForUpLoad) {
        return Tasks.call(mExecutor, () -> {
            File metadata = new File()
                    .setParents(Collections.singletonList("root"))
                    .setMimeType("application/vnd.ms-excel")
                    .setName("Untitled file");
            FileContent mediaContent = new FileContent("application/vnd.ms-excel",fileForUpLoad);
            File googleFile = mDriveService.files().create(metadata,mediaContent).execute();
            if (googleFile == null) {
                throw new IOException("Null result when requesting file creation.");
            }

            return googleFile.getId();
        });
    }

    public Task<Pair<String, String>> readFile(String fileId) {
        return Tasks.call(mExecutor, () -> {
            // Retrieve the metadata as a File object.
            File metadata = mDriveService.files().get(fileId).execute();
            String name = metadata.getName();

            // Stream the file contents to a String.
            try (InputStream is = mDriveService.files().get(fileId).executeMediaAsInputStream();
                 BufferedReader reader = new BufferedReader(new InputStreamReader(is))) {
                StringBuilder stringBuilder = new StringBuilder();
                String line;

                while ((line = reader.readLine()) != null) {
                    stringBuilder.append(line);
                }
                String contents = stringBuilder.toString();

                return Pair.create(name, contents);
            }
        });
    }

}
