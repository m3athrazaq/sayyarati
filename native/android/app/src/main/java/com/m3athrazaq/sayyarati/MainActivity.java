package com.m3athrazaq.sayyarati;

import android.os.Bundle;
import com.getcapacitor.BridgeActivity;

public class MainActivity extends BridgeActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        // the app's own plugins: reading the Excel file the user picks or opens with the app,
        // and keeping the WebView on the app's own pages
        registerPlugin(XlsxImportPlugin.class);
        registerPlugin(NavigationGuardPlugin.class);
        super.onCreate(savedInstanceState);
    }
}
