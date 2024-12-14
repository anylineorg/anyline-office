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
    private String type = "checkbox"; //text:只显示文本
    private String split = ""; //只显示文本时 设置分隔符号

    public void release(){
        super.release();
        data = null;
        valueKey = ConfigTable.DEFAULT_PRIMARY_KEY;
        textKey = "NM";
        property = null;
        rely = null;
        head = null;
        headValue = null;
        value = null;
        checkedValue = "";
        vol = 0;
        label = "";
        checked = false;
        type = "checkbox";
        split = "";
    }
    @Override
    public String parse(String text) {
        StringBuffer html = new StringBuffer();

        if(null == rely) {
            rely = property;
        }
        if(null == rely) {
            rely = valueKey;
        }
        data = fetchAttributeData(text, "data", "d");
        value = fetchAttributeString(text, "value", "v");
        split = fetchAttributeString(text, "split", "s");
        if(null == split){
            split = "";
        }
        if(BasicUtil.isEmpty(data)){
            return "";
        }
        type = fetchAttributeString(text, "type", "t");
        String vk = fetchAttributeString(text, "valueKey", "vk");
        if(BasicUtil.isNotEmpty(vk)){
            valueKey = vk;
        }
        String tk = fetchAttributeString(text, "textKey", "tk");
        if(BasicUtil.isNotEmpty(tk)){
            textKey = tk;
        }
        vol = BasicUtil.parseInt(fetchAttributeString(text, "vol"), vol);

        if (null != data) {
            if (data instanceof String) {
                String items[] = data.toString().split(",");
                List list = new ArrayList();
                for (String item : items) {
                    if(BasicUtil.isEmpty(item)){
                        continue;
                    }
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
            if (null != value) {
                value = context.data(value.toString());
                if (value instanceof Collection) {
                    List list = new ArrayList();
                    Collection cols = (Collection) this.value;
                    for (Object item : cols) {
                        Object val = item;
                        if (item instanceof Map) {
                            val = ((Map) item).get(rely);
                        }
                        list.add(val);
                    }
                    value = list;
                }
            }
            if(!(data instanceof Collection)){
                return "";
            }
            Collection<Map> items = (Collection<Map>) data;
            Collection chks = null;
            if(value instanceof Collection) {
                chks = (Collection<?>) this.value;
            }else{
                chks = new ArrayList<>();
                chks.add(value);
            }


            if (null == headValue) {
                headValue = "";
            }
            int qty = 0;
            if (null != head) {
                qty ++;
                if (null != headValue && !"text".equalsIgnoreCase(type)) {
                    if (checked || checkedValue.equals(headValue) || "true".equalsIgnoreCase(headValue) || "checked".equalsIgnoreCase(headValue) || checked(value, headValue)) {
                        //选中
                        html.append("☑");
                    } else {
                        html.append("☐");
                    }
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
                    if(qty > 1){
                        html.append(split);
                    }
                    if(!"text".equalsIgnoreCase(type)) {
                        if (checkedValue.equals(chk) || "true".equalsIgnoreCase(chk + "") || "checked".equalsIgnoreCase(chk + "") || checked(chks, item.get(valueKey))) {
                            html.append("☑");
                        } else {
                            html.append("☐");
                        }
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
