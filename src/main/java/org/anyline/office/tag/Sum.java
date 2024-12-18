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
import org.anyline.util.BeanUtil;
import org.anyline.util.NumberUtil;

import java.math.BigDecimal;
import java.util.Collection;

public class Sum extends AbstractTag implements Tag{
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
    public void run() throws Exception{
        String html = null;
        property = fetchAttributeString(text, "property", "p");
        format = fetchAttributeString(text, "format","f");
        nvl = fetchAttributeString(text, "nvl", "n");
        min = fetchAttributeString(text, "min");
        max = fetchAttributeString(text, "max");
        def = fetchAttributeString(text, "def");
        scale = BasicUtil.parseInt(fetchAttributeString(text, "scale", "s"), null);
        round = BasicUtil.parseInt(fetchAttributeString(text, "round", "r"), null);
        data = data();
        if(data instanceof Collection){
            Collection items = (Collection) data;
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
        }


        if(BasicUtil.isEmpty(html) && BasicUtil.isNotEmpty(nvl)) {
            html = nvl;
        }
        output(html);
    }

}
