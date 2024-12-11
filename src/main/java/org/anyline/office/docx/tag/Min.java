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

import org.anyline.entity.DataRow;
import org.anyline.entity.DataSet;
import org.anyline.util.BasicUtil;
import org.anyline.util.regular.RegularUtil;

public class Min extends AbstractTag implements Tag{

    public void release(){
        super.release();
    }
    public String parse(String text) throws Exception{
        String html = "";
        String head = RegularUtil.fetchTagHead(text);
        String items_key = fetchAttributeValue(head, "items", "data", "d", "is");
        String property = fetchAttributeValue(head, "property", "p");
        String var = fetchAttributeValue(head, "var", "v");
        Object data = context.data(items_key);
        if(null != data){
            if(data instanceof DataSet){
                DataSet set = (DataSet)data;
                DataRow min = set.min(property);
                if (null != min) {
                    if(BasicUtil.isEmpty(var)) {
                            html = min.getString(property);
                    }else{
                        context.variable(var, min);
                    }
                }
            }
        }
        return html;
    }
}
