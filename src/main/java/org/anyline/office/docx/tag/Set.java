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

import org.anyline.entity.DataSet;
import org.anyline.util.BasicUtil;
import org.anyline.util.BeanUtil;

import java.util.Collection;

public class Set extends AbstractTag implements Tag{

    private String var;
    private Object data;
    private String selector;
    private String distinct;
    private Integer index = null;
    private Integer begin = null;
    private Integer end = null;
    private Integer qty = null;
    public void release(){
        super.release();
        var = null;
        data = null;
        selector = null;
        distinct = null;
        index = null;
        begin = null;
        end = null;
        qty = null;
    }
    public String parse(String text){
        String html = "";
        var = fetchAttributeString(text, "var");
        selector = fetchAttributeString(text, "selector","st");
        distinct = fetchAttributeString(text, "distinct", "ds");
        index = BasicUtil.parseInt(fetchAttributeString(text, "index", "i"), null);
        begin = BasicUtil.parseInt(fetchAttributeString(text, "begin", "b"), null);
        end = BasicUtil.parseInt(fetchAttributeString(text, "end", "e"), null);
        qty = BasicUtil.parseInt(fetchAttributeString(text, "qty", "q"), null);
        data = fetchAttributeData(text, "data", "d");
        if(BasicUtil.isEmpty(data) || BasicUtil.isEmpty(var)){
            return "";
        }
        if (BasicUtil.isNotEmpty(data)) {
            if(data instanceof Collection) {
                Collection items = (Collection) data;
                if(BasicUtil.isNotEmpty(selector)) {
                    items = BeanUtil.select(items,selector.split(","));
                }
                if(index != null) {
                    int i = 0;
                    data = null;
                    for(Object item:items) {
                        if(index ==i) {
                            data = item;
                            break;
                        }
                        i ++;
                    }
                }else{
                    int[] range = BasicUtil.range(begin, end, qty, items.size());
                    if(items instanceof DataSet) {
                        data = ((DataSet) items).cuts(range[0], range[1]);
                    }else {
                        data = BeanUtil.cuts(items, range[0], range[1]);
                    }
                }
                if(null != distinct && data instanceof Collection) {
                    data = BeanUtil.distinct((Collection) data, distinct.split(","));
                }
            }

            doc.variable(var, data);
            doc.replace(var, data.toString());
        }else{
            doc.replace(var, "");
            doc.variable(var, null);
        }
        return html;
    }
}
