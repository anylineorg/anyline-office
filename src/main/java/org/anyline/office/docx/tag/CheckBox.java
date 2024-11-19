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

import java.util.*;

public class CheckBox extends AbstractTag implements Tag {

    private Object data;
    private String valueKey = ConfigTable.DEFAULT_PRIMARY_KEY;
    private String textKey = "NM";
    private String property;
    private String rely;
    private String head;
    private String headValue;
    private String checkedValue = "";
    private boolean checked = false;

    @Override
    public String parse(String text) {
        StringBuffer html = new StringBuffer();
        try {
            if(null == rely) {
                rely = property;
            }
            if(null == rely) {
                rely = valueKey;
            }

            if (null != data) {
                if (data instanceof String) {
                    if (data.toString().endsWith("}")) {
                        data = data.toString().replace("{", "").replace("}", "");
                    } else {
                        data = request.getAttribute(data.toString());
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
                        if(ks.length>2) {
                            map.put("CHK", ks[2]);
                        }
                        list.add(map);
                    }
                    data = list;
                }
                // 选中值
                if (null != this.value) {
                    if(!(this.value instanceof String || this.value instanceof Collection)) {
                        this.value = this.value.toString();
                    }
                    if (this.value instanceof String) {
                        if (this.value.toString().endsWith("}")) {
                            this.value = this.value.toString().replace("{", "").replace("}", "");
                        }
                    }
                    if (this.value instanceof String) {
                        this.value = BeanUtil.array2list(this.value.toString().split(","));
                    }else if(this.value instanceof Collection) {
                        List list = new ArrayList();
                        Collection cols = (Collection)this.value;
                        for(Object item:cols) {
                            Object val = item;
                            if(item instanceof Map) {
                                val = ((Map)item).get(rely);
                            }
                            list.add(val);
                        }
                        this.value = list;
                    }
                }
                Collection<Map> items = (Collection<Map>) data;
                Collection<?> chks = (Collection<?>)this.value;

                // 条目边框
                String itemBorderTagName ="";
                String itemBorderStartTag = "";
                String itemBorderEndTag = "";

                if(BasicUtil.isNotEmpty(border) && !"false".equals(border)) {
                    if("true".equalsIgnoreCase(border)) {
                        itemBorderTagName = "div";
                    }else{
                        itemBorderTagName = border;
                    }
                    itemBorderStartTag = "<"+itemBorderTagName+" class=\""+borderClazz+"\">";
                    itemBorderEndTag = "</"+itemBorderTagName+">";
                }


                if(null == headValue) {
                    headValue = "";
                }
                if(null != head) {
                    String id = this.id;
                    if(BasicUtil.isEmpty(id)) {
                        id = name +"_"+ headValue;
                    }
                    html.append(itemBorderStartTag);
                    html.append("<input type=\"checkbox\"");
                    if(null != headValue) {
                        if(checked || checkedValue.equals(headValue) || "true".equalsIgnoreCase(headValue) || "checked".equalsIgnoreCase(headValue) || checked(value,headValue) ) {
                            html.append(" checked=\"checked\"");
                        }
                    }

                    Map<String, String> map = new HashMap<String, String>();
                    map.put(valueKey, headValue);
                    attribute(html);
                    crateExtraData(html,map);
                    html.append("/>");
                    html.append("<label for=\"").append(id).append("\" class=\""+labelClazz+"\">").append(head).append("</label>\n");

                    html.append(itemBorderEndTag);
                }


                if (null != items)
                    for (Map item : items) {
                        Object val = item.get(valueKey);
                        if(this.encrypt) {
                            val = DESUtil.encryptValue(val+"");
                        }
                        String id = this.id;
                        if(BasicUtil.isEmpty(id)) {
                            id = name +"_"+ val;
                        }
                        html.append(itemBorderStartTag);
                        html.append("<input type=\"checkbox\" value=\"").append(val).append("\" id=\"").append(id).append("\"");
                        Object chk = null;
                        if(BasicUtil.isNotEmpty(rely)) {
                            chk = item.get(rely);
                            if(null != chk) {
                                chk = chk.toString().trim();
                            }
                        }
                        if(checkedValue.equals(chk) || "true".equalsIgnoreCase(chk+"") || "checked".equalsIgnoreCase(chk+"") || checked(chks,item.get(valueKey)) ) {
                            html.append(" checked=\"checked\"");
                        }
                        attribute(html);
                        crateExtraData(html,item);
                        html.append("/>");
                        if(BasicUtil.isEmpty(label)) {
                            String labelHtml = "<label for=\""+id+ "\" class=\""+labelClazz+"\">";
                            String labelBody = "";
                            if (textKey.contains("{")) {
                                labelBody = BeanUtil.parseRuntimeValue(item,textKey);
                            } else {
                                Object v = item.get(textKey);
                                if (null != v) {
                                    labelBody = v.toString();
                                }
                            }
                            labelHtml += labelBody +"</label>\n";
                            html.append(labelHtml);
                        }else{//指定label文本
                            String labelHtml = label;
                            if(labelHtml.contains("{") && labelHtml.contains("}")) {
                                labelHtml = BeanUtil.parseRuntimeValue(item,labelHtml.replace("{","${"));
                            }
                            html.append(labelHtml);
                        }
                        html.append(itemBorderEndTag);
                    }
            }
            return text;
    }
}
