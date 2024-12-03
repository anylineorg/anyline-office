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

package org.anyline.office.docx.tag;

import org.anyline.office.docx.entity.Context;
import org.anyline.office.docx.entity.WDocument;
import org.dom4j.Element;

import java.util.List;

public interface Tag {
    void init(WDocument doc);
    void context(Context context);
    Context context();
    String parse(String text) throws Exception;
    String ref();
    void ref(String ref);
    default Element wt() {
        List<Element> wts = wts();
        if(wts.isEmpty()){
            return null;
        }
        return wts.get(0);
    }
    default void wt(Element wt) {
        wts().add(wt);
    }

    List<Element> wts();
    void wts(List<Element> wts);
    default void release(){}


}
