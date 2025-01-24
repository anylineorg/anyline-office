package org.anyline.office.util;

import java.io.File;

public interface Imager {
    File qr(String content, int width, int height);
    File bar(String content, int width, int height);
}
