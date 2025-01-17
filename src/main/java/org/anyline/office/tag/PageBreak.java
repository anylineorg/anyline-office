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

import org.anyline.util.BasicUtil;

public class PageBreak extends AbstractTag implements Tag {
    public void run() throws Exception{
        String position = fetchAttributeString("position");
        if(BasicUtil.isEmpty(position)){
            position = "after";
        }
        String html = "<div style='page-break-"+position+"'></div>";
        doc.parseHtml(tops.get(0), contents.get(0), html);
        box.remove();
    }
}
