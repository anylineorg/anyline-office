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

import org.anyline.office.docx.util.DocxUtil;
import org.anyline.util.BasicUtil;
import org.anyline.util.DomUtil;
import org.anyline.util.NumberUtil;
import org.dom4j.Element;

import java.util.List;

public class Mark extends AbstractTag implements Tag{

    public void release(){
        super.release();
    }
    public void run(){
        Element first = contents.get(0);
        Element last = contents.get(contents.size()-1);
        Element parent = first.getParent();
        int max = -1;
        List<Element> starts = DomUtil.elements(true, doc.getSrc(), "bookmarkStart");
        for (Element start : starts){
            String value = start.attributeValue("w:id");
            int id = BasicUtil.parseInt(value, 0);
            max = NumberUtil.max(id, max);
        }
        max ++;
        String name = fetchAttributeString("name", "n");
        Element start = DocxUtil.addElement(parent, "bookmarkStart");
        Element end = DocxUtil.addElement(parent, "bookmarkEnd");
        start.addAttribute("w:id", max+"");
        start.addAttribute("w:name", name);;
        end.addAttribute("w:id", max+"");
        DocxUtil.before(start, first);
        DocxUtil.after(end, last);
        DocxUtil.remove(first);
        DocxUtil.remove(last);
    }
}
