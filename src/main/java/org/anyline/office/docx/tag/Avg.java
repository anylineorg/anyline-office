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
import org.anyline.util.regular.RegularUtil;

public class Avg extends AbstractTag implements Tag {
    private Object data;
    private String property;
    private String var;
    private String distinct;
    private int scale = 2;
    private int round = 4;

    public void release(){
        super.release();
        property = null;
        data = null;
        var = null;
        distinct = null;
        scale = 2;
        round = 4;
    }
    public String parse(String text) {
        String result = "";
        try {
            String key = RegularUtil.fetchAttributeValue(text, "data");
            property = RegularUtil.fetchAttributeValue(text, "property");
            var = RegularUtil.fetchAttributeValue(text, "var");
            distinct = RegularUtil.fetchAttributeValue(text, "distinct");
            scale = BasicUtil.parseInt(RegularUtil.fetchAttributeValue(text, "scale"), scale);
            round = BasicUtil.parseInt(RegularUtil.fetchAttributeValue(text, "round"), round);

            if(BasicUtil.isEmpty(key)){
                return "";
            }
            data = context.data(key);

            if(data instanceof String) {
                String[] items = data.toString().split(",");
            }else if(data instanceof DataSet){
                DataSet set = (DataSet) data;
                if(BasicUtil.isNotEmpty(distinct)){
                    set = set.distinct(distinct.split(","));
                }
                if(null != property){
                    result = set.avg(scale, round, property.split(","))+"";
                }

            }
            if(BasicUtil.isNotEmpty(var)){
                doc.context().variable(var, result);
                result = "";
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }
}
