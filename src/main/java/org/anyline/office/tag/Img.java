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

public class Img extends AbstractTag implements Tag{
    public void release(){
        super.release();
    }
    @Override
    public void run() {
        String result = context.placeholder(text);
        //<aot:img src=”${FILE_URL_COL}” style=”width:150px;height:${LOGO_HEIGHT}px;”></aot:img>
        result = result.replace(doc.namespace() +":img", "img");
        result = context.placeholder(result);
        doc.parseHtml(tops.get(0), contents.get(0), result);
        //String placeholder = "__"+BasicUtil.getRandomString(16); //__开头的占位符 没有值时中保留原样 有可能是下一步需要处理的

        //doc.variable(placeholder, result);
        //output("${"+placeholder+"}");
        box.remove();
    }
}
