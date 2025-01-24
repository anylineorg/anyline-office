/*
 * Copyright 2006-2023 www.anyline.org
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.anyline.office.tag;

import org.anyline.office.util.Imager;
import org.anyline.util.BasicUtil;
import org.anyline.util.StyleParser;

import javax.imageio.ImageReader;
import java.io.File;
import java.util.Map;

public class QRCode extends AbstractTag implements Tag{
    public void release(){
        super.release();
    }
    @Override
    public void run() {
        //<aot:img src=”${FILE_URL_COL}” style=”width:150px;height:${LOGO_HEIGHT}px;”></aot:img>
        String value = fetchAttributeString("value", "v");
        if(BasicUtil.isEmpty(value)){
            value = body(text, "qr");
        }
        String style = fetchAttributeString("style");
        Map<String, String> styles = StyleParser.parse(style);
        String width = styles.get("width");
        String height = styles.get("height");
        if(BasicUtil.isEmpty(width)){
            width = fetchAttributeString("width", "w");
        }
        if(BasicUtil.isEmpty(height)){
            height = fetchAttributeString("height", "h");
        }
        int w = 100;
        int h = 100;
        if(null != width){
            w = BasicUtil.parseInt(width.replace("px", "").trim(), w);
        }
        if(null != height){
            h = BasicUtil.parseInt(height.replace("px", "").trim(), h);
        }

        Imager imager = doc.getImager();
        if(null != imager) {
            File file = imager.qr(value, w, h);
            String result = "<img src='file:" + file.getAbsolutePath() + "' style='" + style + "'/>";
            result = context.placeholder(result);
            doc.parseHtml(tops.get(0), contents.get(0), result);
            file.delete();
        }
        box.remove();
    }
}
