package com.m3athrazaq.sayyarati;

import android.net.Uri;
import com.getcapacitor.Plugin;
import com.getcapacitor.annotation.CapacitorPlugin;

/**
 * The app never leaves its own pages. Capacitor would hand any other address to another app (the
 * browser); here every navigation outside https://localhost is blocked instead, so even a future bug in
 * the page could not use a link to carry data out of the app.
 */
@CapacitorPlugin(name = "NavigationGuard")
public class NavigationGuardPlugin extends Plugin {

    @Override
    public Boolean shouldOverrideLoad(Uri url) {
        if (url != null && "https".equals(url.getScheme()) && "localhost".equals(url.getHost())) {
            return null; // the app's own pages: normal handling
        }
        return true; // anything else: do nothing
    }
}
