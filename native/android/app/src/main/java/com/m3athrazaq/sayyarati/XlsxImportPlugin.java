package com.m3athrazaq.sayyarati;

import android.app.Activity;
import android.content.ContentResolver;
import android.content.Context;
import android.content.Intent;
import android.content.pm.ApplicationInfo;
import android.content.res.AssetFileDescriptor;
import android.database.Cursor;
import android.net.Uri;
import android.provider.DocumentsContract;
import android.provider.OpenableColumns;
import androidx.activity.result.ActivityResult;
import androidx.core.content.IntentCompat;
import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.ActivityCallback;
import com.getcapacitor.annotation.CapacitorPlugin;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InterruptedIOException;
import java.io.OutputStream;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * Brings an Excel file into the app.
 *
 * The WebView's own file input reads the chosen file lazily, which fails for Google Drive files that
 * are not downloaded yet. Here the file is read completely through the ContentResolver instead (Drive
 * downloads it while we read), and a Google Sheets document is exported by Drive as .xlsx.
 *
 * - pick(): opens the system file picker (Drive, Downloads, Files...).
 * - takeOpenedFile(): the file the app was opened with ("Open with" / "Share" from another app).
 *
 * Only content:// files the user chose are read, at most MAX_BYTES, on the plugin's own threads with a
 * time limit (a provider that never answers cannot block the rest of the app). The file is copied into
 * the app's own cache (cache/imports/, which nothing else can read); the page reads that copy, deletes
 * it, and validates the content before anything is imported. A backup with invoice photos can be large,
 * so the bytes are never passed to the page as one huge message.
 */
@CapacitorPlugin(name = "XlsxImport")
public class XlsxImportPlugin extends Plugin {

    static final String XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    static final long MAX_BYTES = 150L * 1024 * 1024;
    static final long TIMEOUT_SECONDS = 300;
    static final String IMPORT_DIR = "imports";

    private final ExecutorService reader = Executors.newFixedThreadPool(2, r -> daemon(r, "XlsxImport-read"));
    private final ScheduledExecutorService timer = Executors.newSingleThreadScheduledExecutor(r -> daemon(r, "XlsxImport-timer"));

    private static Thread daemon(Runnable r, String name) {
        Thread t = new Thread(r, name);
        t.setDaemon(true);
        return t;
    }

    /** file handed to the app by another app, waiting for the page to take it */
    private Uri pendingUri;

    @Override
    protected void handleOnNewIntent(Intent intent) {
        super.handleOnNewIntent(intent);
        Uri uri = incomingUri(intent);
        if (uri != null && allowed(uri)) {
            pendingUri = uri;
            notifyListeners("fileOpened", new JSObject(), true);
        }
    }

    /** VIEW (Open with) or SEND (Share) of a content:// file; anything else is ignored */
    static Uri incomingUri(Intent intent) {
        if (intent == null) return null;
        String action = intent.getAction();
        Uri uri = null;
        if (Intent.ACTION_VIEW.equals(action)) {
            uri = intent.getData();
        } else if (Intent.ACTION_SEND.equals(action)) {
            uri = IntentCompat.getParcelableExtra(intent, Intent.EXTRA_STREAM, Uri.class);
            if (uri == null && intent.getClipData() != null && intent.getClipData().getItemCount() > 0) {
                uri = intent.getClipData().getItemAt(0).getUri();
            }
        }
        if (uri == null || !ContentResolver.SCHEME_CONTENT.equals(uri.getScheme())) return null;
        return uri;
    }

    /**
     * Files from other apps only: an address of this app's own FileProvider is refused, so another app
     * cannot make it reopen its own cached files. (The emulator test hands its sample files over that
     * way, so debug builds allow it.)
     */
    private boolean allowed(Uri uri) {
        String authority = uri.getAuthority();
        if (authority == null || authority.isEmpty()) return false;
        if (!authority.startsWith(getContext().getPackageName())) return true;
        return (getContext().getApplicationInfo().flags & ApplicationInfo.FLAG_DEBUGGABLE) != 0;
    }

    @PluginMethod
    public void takeOpenedFile(PluginCall call) {
        Uri uri = pendingUri;
        pendingUri = null;
        if (uri == null) {
            JSObject r = new JSObject();
            r.put("none", true);
            call.resolve(r);
            return;
        }
        // the same launch intent must not be imported again if the activity is recreated
        Activity activity = getActivity();
        if (activity != null) activity.runOnUiThread(() -> activity.setIntent(new Intent(Intent.ACTION_MAIN)));
        readInBackground(call, uri, false);
    }

    @PluginMethod
    public void pick(PluginCall call) {
        Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
        // No CATEGORY_OPENABLE, so Google Sheets documents in Drive are offered too (read below as .xlsx).
        intent.setType("*/*");
        intent.putExtra(
            Intent.EXTRA_MIME_TYPES,
            // some apps label .xlsx files loosely; the page checks the content before importing
            new String[] { XLSX, "application/vnd.google-apps.spreadsheet", "application/vnd.ms-excel", "application/octet-stream" }
        );
        startActivityForResult(call, intent, "picked");
    }

    @ActivityCallback
    private void picked(PluginCall call, ActivityResult result) {
        if (call == null) return;
        Intent data = result.getData();
        if (result.getResultCode() != Activity.RESULT_OK || data == null || data.getData() == null) {
            JSObject r = new JSObject();
            r.put("cancelled", true);
            call.resolve(r);
            getBridge().releaseCall(call);
            return;
        }
        readInBackground(call, data.getData(), true);
    }

    private void readInBackground(PluginCall call, Uri uri, boolean saved) {
        AtomicBoolean finished = new AtomicBoolean(false);
        Future<?> job = reader.submit(() -> {
            JSObject r = null;
            String error = null, code = null;
            File copy = null;
            try {
                Context ctx = getContext();
                ContentResolver cr = ctx.getContentResolver();
                File dir = new File(ctx.getCacheDir(), IMPORT_DIR);
                clearDir(dir);
                if (!dir.isDirectory() && !dir.mkdirs()) throw new IOException("cannot create the import folder");
                copy = new File(dir, "import-" + System.currentTimeMillis() + ".xlsx");
                long size = copyTo(ctx, cr, uri, copy);
                r = new JSObject();
                r.put("name", displayName(cr, uri));
                r.put("size", size);
                r.put("path", IMPORT_DIR + "/" + copy.getName());
                r.put("uri", Uri.fromFile(copy).toString());
            } catch (TooLargeException e) {
                error = "The file is larger than " + (MAX_BYTES / 1024 / 1024) + " MB";
                code = "TOO_LARGE";
            } catch (NoXlsxException e) {
                error = "This Google Sheets file cannot be exported as .xlsx here";
                code = "NO_XLSX";
            } catch (SecurityException e) {
                error = "No permission to read the file";
                code = "NO_PERMISSION";
            } catch (Exception e) {
                error = e.getClass().getSimpleName() + ": " + e.getMessage();
                code = "READ_FAILED";
            }
            if (r == null && copy != null) copy.delete();
            if (!finished.compareAndSet(false, true)) { // timed out already
                if (copy != null) copy.delete();
                return;
            }
            if (r != null) call.resolve(r);
            else call.reject(error, code);
            if (saved) getBridge().releaseCall(call);
        });
        timer.schedule(
            () -> {
                if (!finished.compareAndSet(false, true)) return;
                job.cancel(true);
                call.reject("Reading the file took too long", "TIMEOUT");
                if (saved) getBridge().releaseCall(call);
            },
            TIMEOUT_SECONDS,
            TimeUnit.SECONDS
        );
    }

    /** removes copies left behind (the page reads an import and deletes it right away) */
    private static void clearDir(File dir) {
        File[] old = dir.listFiles();
        if (old == null) return;
        long cutoff = System.currentTimeMillis() - 10 * 60 * 1000;
        for (File f : old) if (f.isFile() && f.lastModified() < cutoff) f.delete();
    }

    /** copies the file into `out`; stops at MAX_BYTES, or when the time limit cancels the job */
    private static long copyTo(Context ctx, ContentResolver cr, Uri uri, File out) throws IOException {
        InputStream in = null;
        AssetFileDescriptor afd = null;
        try {
            if (isVirtual(ctx, cr, uri)) {
                // e.g. a Google Sheets document: ask the provider for an .xlsx copy
                String[] types = cr.getStreamTypes(uri, XLSX);
                if (types == null || types.length == 0) throw new NoXlsxException();
                afd = cr.openTypedAssetFileDescriptor(uri, XLSX, null);
                if (afd == null) throw new NoXlsxException();
                in = afd.createInputStream();
            } else {
                in = cr.openInputStream(uri);
            }
            if (in == null) throw new IOException("no data");
            long total = 0;
            try (OutputStream os = new FileOutputStream(out)) {
                byte[] buf = new byte[256 * 1024];
                int n;
                while ((n = in.read(buf)) != -1) {
                    if (Thread.currentThread().isInterrupted()) throw new InterruptedIOException("cancelled");
                    total += n;
                    if (total > MAX_BYTES) throw new TooLargeException();
                    os.write(buf, 0, n);
                }
            }
            return total;
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException ignored) {}
            }
            if (afd != null) {
                try {
                    afd.close();
                } catch (IOException ignored) {}
            }
        }
    }

    private static boolean isVirtual(Context ctx, ContentResolver cr, Uri uri) {
        if (!DocumentsContract.isDocumentUri(ctx, uri)) return false;
        try (Cursor c = cr.query(uri, new String[] { DocumentsContract.Document.COLUMN_FLAGS }, null, null, null)) {
            if (c != null && c.moveToFirst() && !c.isNull(0)) {
                return (c.getInt(0) & DocumentsContract.Document.FLAG_VIRTUAL_DOCUMENT) != 0;
            }
        } catch (Exception ignored) {}
        return false;
    }

    private static String displayName(ContentResolver cr, Uri uri) {
        String name = null;
        try (Cursor c = cr.query(uri, new String[] { OpenableColumns.DISPLAY_NAME }, null, null, null)) {
            if (c != null && c.moveToFirst()) {
                int i = c.getColumnIndex(OpenableColumns.DISPLAY_NAME);
                if (i >= 0) name = c.getString(i);
            }
        } catch (Exception ignored) {}
        if (name == null) name = uri.getLastPathSegment();
        if (name == null || name.isEmpty()) name = "file.xlsx";
        return name.length() > 120 ? name.substring(0, 120) : name;
    }

    static class TooLargeException extends IOException {}

    static class NoXlsxException extends IOException {}
}
