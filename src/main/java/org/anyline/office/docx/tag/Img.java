package org.anyline.office.docx.tag;

import org.anyline.util.BasicUtil;

public class Img extends AbstractTag implements Tag{
    @Override
    public String parse(String text) {
        String result = placeholder(text);
        //<aol:img src=”${FILE_URL_COL}” style=”width:150px;height:${LOGO_HEIGHT}px;”></aol:img>
        result = result.replace("aol:img", "img");
        String placeholder = BasicUtil.getRandomString(16);
        doc.replace(placeholder, result);
        return "${"+placeholder+"}";
    }
}
