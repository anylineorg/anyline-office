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
import org.anyline.util.NumberUtil;
import org.anyline.util.regular.RegularUtil;

import java.math.BigDecimal;
import java.util.Collection;

public class Sum extends AbstractTag implements Tag{

    private String var;
    private Object data;
    //选择过滤器  ID:1,TYPE:2
    private String selector;
    private String property;
    private String format;
    private String nvl;
    private Object min;
    private Object max;
    private String def; // 默认值
    private Integer scale;//小数位
    private Integer round; // 参考BigDecimal.ROUND_UP;
    public void release(){
        super.release();
        var = null;
        data = null;
        selector = null;
        property = null;
        format = null;
        nvl = null;
        min = null;
        max = null;
        def = null;
        scale = null;
        round = null;
    }
    public String parse(String text){
        String html = "";
        String key = RegularUtil.fetchAttributeValue(text, "data");
        var = RegularUtil.fetchAttributeValue(text, "var");
        selector = RegularUtil.fetchAttributeValue(text, "selector");
        property = RegularUtil.fetchAttributeValue(text, "property");
        format = RegularUtil.fetchAttributeValue(text, "format");
        nvl = RegularUtil.fetchAttributeValue(text, "nvl");
        min = RegularUtil.fetchAttributeValue(text, "min");
        max = RegularUtil.fetchAttributeValue(text, "max");
        def = RegularUtil.fetchAttributeValue(text, "def");
        scale = BasicUtil.parseInt(RegularUtil.fetchAttributeValue(text, "selector"), null);
        round = BasicUtil.parseInt(RegularUtil.fetchAttributeValue(text, "round"), null);
        if(BasicUtil.isEmpty(key)){
            return html;
        }
        data = context.data(key);
        if(BasicUtil.isEmpty(data)){
            return html;
        }

        Collection items = (Collection) data;
        if(BasicUtil.isNotEmpty(selector) && data instanceof DataSet) {
            items = BeanUtil.select(items,selector.split(","));
        }

        BigDecimal sum = new BigDecimal(0);
        if (null != items) {
            for (Object item : items) {
                if(null == item) {
                    continue;
                }
                Object val = null;
                if(item instanceof Number) {
                    val = item;
                }else{
                    val = BeanUtil.getFieldValue(item, property);
                }
                if(null != val) {
                    sum = sum.add(new BigDecimal(val.toString()));
                }
            }

            if(BasicUtil.isNotEmpty(min)) {
                BigDecimal minNum = new BigDecimal(min.toString());
                if(minNum.compareTo(sum) > 0) {
                    sum = minNum;
                }
            }
            if(BasicUtil.isNotEmpty(max)) {
                BigDecimal maxNum = new BigDecimal(max.toString());
                if(maxNum.compareTo(sum) < 0) {
                    sum = maxNum;
                }
            }
            if(null != scale) {
                if(null != round) {
                    sum = sum.setScale(scale, round);
                }else {
                    sum = sum.setScale(scale);
                }
            }
            if(BasicUtil.isNotEmpty(format)) {
                html = NumberUtil.format(sum,format);
            }else{
                html = sum.toString();
            }
        }

        if(BasicUtil.isEmpty(html) && BasicUtil.isNotEmpty(nvl)) {
            html = nvl;
        }
        if(null != var){
            doc.context().variable(var, html);
            html = "";
        }

        return html;
    }

}
