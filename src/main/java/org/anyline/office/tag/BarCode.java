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

import java.io.File;

public class BarCode extends AbstractTag implements Tag{
    public void release(){
        super.release();
    }
    @Override
    public void run() {
        String value = fetchAttributeString("value", "v");
        if(BasicUtil.isEmpty(value)){
            value = body(text, "bar");
        }
        int width = BasicUtil.parseInt(fetchAttributeString("width", "w"), 100);
        int height = BasicUtil.parseInt(fetchAttributeString("height", "h"), 100);
        String style = "width:"+width+"px;height:"+height+"px;";
        Imager imager = doc.getImager();
        if(null != imager) {
            File file = imager.bar(value, width, height);
            String result = "<img src='file:" + file.getAbsolutePath() + "' style='" + style + "'/>";
            result = context.placeholder(result);
            doc.parseHtml(tops.get(0), contents.get(0), result);
        }
        box.remove();
    }
}
