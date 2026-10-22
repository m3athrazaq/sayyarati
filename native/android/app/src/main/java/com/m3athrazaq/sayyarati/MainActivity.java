package com.m3athrazaq.sayyarati;

import android.os.Bundle;
import com.getcapacitor.BridgeActivity;

public class MainActivity extends BridgeActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        // the app's own plugin (reads the Excel file the user picks or opens with the app)
        registerPlugin(XlsxImportPlugin.class);
        super.onCreate(savedInstanceState);
    }
}
