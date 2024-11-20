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

import org.anyline.util.BasicUtil;
import org.anyline.util.BeanUtil;
import org.anyline.util.ConfigTable;
import org.anyline.util.encrypt.DESUtil;
import org.anyline.util.regular.RegularUtil;

import java.util.*;

public class CheckBox extends AbstractTag implements Tag {

    private Object data;
    private String valueKey = ConfigTable.DEFAULT_PRIMARY_KEY;
    private String textKey = "NM";
    private String property;
    /**
     * 根据哪个属性判断选中
     */
    private String rely;
    /**
     * 默认选项
     */
    private String head;
    private String headValue;
    /**
    * 当前值
    * */
    protected Object value;
    /**
     * 选中值(checkValue == value时选中)
     */
    private String checkedValue = "";
    /**
     * 每行选项数量
     * 0表示一行显示所有
     */
    private int vol = 0;
    private String label = "";//label标签体,如果未定义label则生成默认label标签体{textKey}

    /**
     * 是否选中
     * 一般在单选项时用到
     */
    private boolean checked = false;

    @Override
    public String parse(String text) {
        StringBuffer html = new StringBuffer();

        if(null == rely) {
            rely = property;
        }
        if(null == rely) {
            rely = valueKey;
        }
        data = RegularUtil.fetchAttributeValue(text, "data");
        value = RegularUtil.fetchAttributeValue(text, "value");
        String vk = RegularUtil.fetchAttributeValue(text, "valueKey");
        if(BasicUtil.isNotEmpty(vk)){
            valueKey = vk;
        }
        String tk = RegularUtil.fetchAttributeValue(text, "textKey");
        if(BasicUtil.isNotEmpty(tk)){
            textKey = tk;
        }
        vol = BasicUtil.parseInt(RegularUtil.fetchAttributeValue(text, "vol"), vol);

        if (null != data) {
            if (data instanceof String) {
                String str = (String)data;
                if (str.startsWith("{") && str.endsWith("}")) {
                    //{0:否,1:是}
                    data = str.replace("{", "").replace("}", "");
                } else if(BasicUtil.checkEl(str)){
                    String key = str.substring(2, str.length() - 1);
                    data = variables.get(key);
                }
            }
            if (data instanceof String) {
                String items[] = data.toString().split(",");
                List list = new ArrayList();
                for (String item : items) {
                    Map map = new HashMap();
                    String ks[] = BeanUtil.parseKeyValue(item);
                    map.put(valueKey, ks[0]);
                    map.put(textKey, ks[1]);
                    if (ks.length > 2) {
                        map.put("CHK", ks[2]);
                    }
                    list.add(map);
                }
                data = list;
            }
            // 选中值
            if (null != this.value) {
                if (!(this.value instanceof String || this.value instanceof Collection)) {
                    this.value = this.value.toString();
                }
                if (this.value instanceof String) {
                    String str = (String) value;
                    if (str.startsWith("{") && str.endsWith("}")) {
                        value = str.replace("{", "").replace("}", "");
                    }
                }
                if (this.value instanceof String) {
                    this.value = BeanUtil.array2list(this.value.toString().split(","));
                } else if (this.value instanceof Collection) {
                    List list = new ArrayList();
                    Collection cols = (Collection) this.value;
                    for (Object item : cols) {
                        Object val = item;
                        if (item instanceof Map) {
                            val = ((Map) item).get(rely);
                        }
                        list.add(val);
                    }
                    this.value = list;
                }
            }
            Collection<Map> items = (Collection<Map>) data;
            Collection<?> chks = (Collection<?>) this.value;


            if (null == headValue) {
                headValue = "";
            }
            int qty = 0;
            if (null != head) {
                if (null != headValue) {
                    if (checked || checkedValue.equals(headValue) || "true".equalsIgnoreCase(headValue) || "checked".equalsIgnoreCase(headValue) || checked(value, headValue)) {
                        //选中
                        html.append("☑");
                    } else {
                        html.append("☐");
                    }
                    qty ++;
                }
                html.append(head);
                if(vol > 0 && qty%vol == 0){
                    html.append("<br/>");
                }
            }
            if (null != items){
                int size = items.size();
                int idx = 0;
                for (Map item : items) {
                    idx ++;
                    qty ++;
                    Object val = item.get(valueKey);
                    Object chk = null;
                    if(BasicUtil.isNotEmpty(rely)) {
                        chk = item.get(rely);
                        if(null != chk) {
                            chk = chk.toString().trim();
                        }
                    }
                    if(checkedValue.equals(chk) || "true".equalsIgnoreCase(chk+"") || "checked".equalsIgnoreCase(chk+"") || checked(chks,item.get(valueKey)) ) {
                        html.append("☑");
                    }else{
                        html.append("☐");
                    }
                    if(BasicUtil.isEmpty(label)) {
                        String labelBody = "";
                        if (textKey.contains("{")) {
                            labelBody = BeanUtil.parseRuntimeValue(item,textKey);
                        } else {
                            Object v = item.get(textKey);
                            if (null != v) {
                                labelBody = v.toString();
                            }
                        }
                        html.append(labelBody);
                    }else{//指定label文本
                        String labelHtml = label;
                        if(labelHtml.contains("{") && labelHtml.contains("}")) {
                            labelHtml = BeanUtil.parseRuntimeValue(item,labelHtml.replace("{","${"));
                        }
                        html.append(labelHtml);
                    }
                    if(vol > 0 && qty%vol == 0 && idx<size){
                        html.append("<br/>");
                    }
                }
            }
        }

        String placeholder = BasicUtil.getRandomString(16);
        doc.replace(placeholder, html.toString());
        return "${"+placeholder+"}";
    }
    private boolean checked(Collection<?> chks, Object value) {
        if(null != chks) {
            for(Object chk:chks) {
                if(null != chk && null != value && chk.toString().equals(value.toString())) {
                    return true;
                }
            }
        }
        return false;
    }
    @SuppressWarnings("rawtypes")
    private boolean checked(Object chks, Object value) {
        if(null != chks) {
            if(chks instanceof Collection) {
                return checked((Collection)chks, value);
            }else{
                return chks.equals(value);
            }
        }
        return false;
    }
}
