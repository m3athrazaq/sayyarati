package com.m3athrazaq.sayyarati;

import android.os.Bundle;
import android.os.Environment;
import android.os.Handler;
import android.os.Looper;
import com.getcapacitor.BridgeActivity;
import java.io.File;

public class MainActivity extends BridgeActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        // the app's own plugins: reading the Excel file the user picks or opens with the app,
        // and keeping the WebView on the app's own pages
        registerPlugin(XlsxImportPlugin.class);
        registerPlugin(NavigationGuardPlugin.class);
        super.onCreate(savedInstanceState);
    }

    private final Handler handler = new Handler(Looper.getMainLooper());
    private final Runnable lateCleanup = () -> cleanCaptures(45 * 1000);

    @Override
    public void onResume() {
        super.onResume();
        // coming back from the camera app: the page reads the new picture within seconds, so a minute
        // later every temporary picture older than 45 seconds can go
        cleanCaptures(10 * 60 * 1000);
        handler.removeCallbacks(lateCleanup);
        handler.postDelayed(lateCleanup, 60 * 1000);
    }

    @Override
    public void onDestroy() {
        handler.removeCallbacks(lateCleanup);
        super.onDestroy();
    }

    /**
     * "Take a photo" of an invoice: the camera app writes the full-size picture to a temporary file in
     * the app's own Pictures folder, and the page saves a smaller copy of it (without the camera's hidden
     * data). The full-size temporary files are deleted here once older than `ageMillis`.
     */
    private void cleanCaptures(long ageMillis) {
        final File dir = getExternalFilesDir(Environment.DIRECTORY_PICTURES);
        if (dir == null) return;
        new Thread(() -> {
            File[] files = dir.listFiles();
            if (files == null) return;
            long cutoff = System.currentTimeMillis() - ageMillis;
            for (File f : files) {
                if (f.isFile() && f.getName().startsWith("JPEG_") && f.lastModified() < cutoff) f.delete();
            }
        }, "captures-cleanup").start();
    }
}
