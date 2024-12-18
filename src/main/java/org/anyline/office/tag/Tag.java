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

import org.anyline.office.docx.entity.WDocument;
import org.anyline.office.util.Context;
import org.dom4j.Element;

import java.util.List;

public interface Tag {
    void init(WDocument doc);

    /**
     * parse前准备工作
     */
    void prepare();
    void context(Context context);
    Context context();
    String run() throws Exception;
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

    /**
     * 标签内的wt所在的顶层p或table
     * 注意如果是与标签在同一个wp中的 设置top=wt
     * @return list
     */
    List<Element> tops();
    void tops(List<Element> tops);
    default void release(){}
    void text(String text);
    String text();


}
