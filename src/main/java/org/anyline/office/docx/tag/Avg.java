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
import org.anyline.util.NumberUtil;

import java.math.BigDecimal;

public class Avg extends AbstractTag implements Tag {
    private Object data;
    private String property;
    private String var;
    private String distinct;
    private int scale = 2;
    private int round = 4;
    private String format;

    public void release(){
        super.release();
        property = null;
        data = null;
        var = null;
        distinct = null;
        scale = 2;
        round = 4;
        format = null;
    }
    public String parse(String text) throws Exception{
        String result = "";
        property = fetchAttributeString(text, "property", "p");
        var = fetchAttributeString(text, "var");
        distinct = fetchAttributeString(text, "distinct", "ds");
        scale = BasicUtil.parseInt(fetchAttributeString(text, "scale", "s"), scale);
        round = BasicUtil.parseInt(fetchAttributeString(text, "round", "r"), round);

        data = fetchAttributeData(text, "data", "d");
        if(data instanceof String) {
            //TODO
            String[] items = data.toString().split(",");
        }else if(data instanceof DataSet){
            BigDecimal avg = null;
            DataSet set = (DataSet) data;
            if(BasicUtil.isNotEmpty(distinct)){
                set = set.distinct(distinct.split(","));
            }
            if(null != property){
                avg = set.avg(scale, round, property.split(","));
            }
            if(null != avg){
                if(null != format){
                    result = NumberUtil.format(avg, format);
                }else{
                    result = avg.toString();
                }
            }
        }
        if(BasicUtil.isNotEmpty(var)){
            doc.context().variable(var, result);
            result = "";
        }
        return result;
    }
}
