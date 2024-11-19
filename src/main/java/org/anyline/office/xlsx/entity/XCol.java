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

package org.anyline.office.xlsx.entity;

import org.anyline.office.docx.util.DocxUtil;
import org.anyline.util.BasicUtil;
import org.dom4j.Element;

import java.util.*;

/**
 * 参考 https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cell?view=openxml-3.0.1
 * 说明 http://office.anyline.org/v/86_14141
 */
public class XCol extends XElement{
    private XRow row;
    private String r        ; // A1
    private String type     ; // t属性
    private String style    ; // s属性
    private String value    ; // ShareString.id或text t="s"时 value=ShareString
    private String text     ; // 最终文本
    private String formula  ; // 公式
    private int x = 0       ; // 行号从1开始
    private String y        ; // 列行从A开始
    private int index       ; //第几列 从0开始

    public XCol(XWorkBook book, XSheet sheet, XRow row, Element src, int index){
        this.book = book;
        this.sheet = sheet;
        this.row = row;
        this.src = src;
        this.index = index;
        this.x = row.r();
        load();
    }
    public void load(){
        if(null == src){
            return;
        }
        type = src.attributeValue("t");
        style = src.attributeValue("s");
        Element ev = src.element("v");
        if(null != ev){
            value = ev.getTextTrim();
        }
        r = src.attributeValue("r");
        y = r.replaceAll("\\d+", "");
    }
    public String r(){
        return r;
    }
    public XCol r(String r){
        this.r = r;
        src.attribute("r").setValue(r);
        return this;
    }
    public int x(){
        return x;
    }
    public XCol x(int x){
        this.x = x;
        return this;
    }

    public String y(){
        return y;
    }
    public XCol y(String y){
        this.x = x;
        return this;
    }

    public List<String> placeholders(String regex){
        List<String> placeholders = new ArrayList<>();
        if("s".equals(type)){
            //文本类型
            int idx = Integer.parseInt(value);
            ShareString ss = book.share(idx);
            if(null != ss) {
                String txt = ss.text();
                if(null != txt) {
                    List<String> flags = DocxUtil.splitKey(txt, regex);
                    if(!flags.isEmpty()){
                        for(String flag:flags){
                            String key = null;
                            if(flag.startsWith("${") && flag.endsWith("}")) {
                                key = flag.substring(2, flag.length() - 1);
                                placeholders.add(key);
                            }
                        }
                    }
                }
            }
        }
        return placeholders;
    }
    public List<String> placeholders(){
        return placeholders("\\$\\{.*?\\}");
    }
    /**
     * 计算列号
     * @param index 下标从0开始 0=A 25=Z 26=AA
     * @return String
     */
    public static String y(int index){
        StringBuilder builder = new StringBuilder();
        int remainder = index + 1; // Excel列是从1开始的，因此需要+1
        while (remainder > 0) {
            remainder--; // 转换为从0开始的索引，方便计算
            int modulo = remainder % 26;
            builder.insert(0, (char)('A' + modulo));
            remainder /= 26;
        }
        return builder.toString();
    }


    public int index(){
        return index;
    }
    public XCol index(int index){
        this.index = index;
        return this;
    }
    public static XCol build(XWorkBook book, XSheet sheet, XRow xr, XCol template, Object value, int x, int y){
        Element row = xr.getSrc();
        Element c = row.addElement("c");
        int r = xr.r(); //行号
        c.addAttribute("s", template.style);
        c.addAttribute("r", y(y)+r);
        Element v = c.addElement("v");
        if(BasicUtil.isNotEmpty(value)){
            if(BasicUtil.isNumber(value)){
                v.setText(value.toString());
            }else{
                int share = book.share(value.toString());
                c.addAttribute("t","s");
                v.setText(share+"");
            }
        }
        XCol xc = new XCol(book, sheet, xr, c, y);
        return xc;
    }
    /**
     * 解析标签
     */
    public void parseTag(){
    }
    public void replace(boolean parse, LinkedHashMap<String, String> replaces){
        if("s".equals(type)){
            //文本类型
            int idx = Integer.parseInt(value);
            ShareString ss = book.share(idx);
            if(null != ss) {
                String txt = ss.text();
                if(null != txt) {
                    String result = replace(parse, txt, replaces);
                    if(!txt.equals(result)) {
                        Element v = src.element("v");
                        if(BasicUtil.isEmpty(result)){
                            //没有内容的清空v标签
                            if(null != v){
                                src.remove(v);
                            }
                        }else{
                            int index = book.share(result);
                            //添加新ShareString并引用
                            if(null == v){
                                v = src.addElement("v");
                            }
                            v.setText(index+"");
                            value = index+"";
                        }
                    }
                }
            }
        }
    }

    public static String replace(boolean parse, String src, Map<String, String> replaces){
        String txt = src;
        List<String> flags = DocxUtil.splitKey(txt);
        if(flags.size() == 0){
            return src;
        }
        Collections.reverse(flags);
        boolean exists = false;
        for(int i=0; i<flags.size(); i++){
            String flag = flags.get(i);
            String content = flag;
            String key = null;
            if(flag.startsWith("${") && flag.endsWith("}")) {
                key = flag.substring(2, flag.length() - 1);
                content = replaces.get(key);
                exists = exists || replaces.containsKey(key);
                if(null == content){
                    exists =  exists || replaces.containsKey(flag);
                    content = replaces.get(flag);
                }
            }else if(flag.startsWith("{") && flag.endsWith("}")){
                key = flag.substring(1, flag.length() - 1);
                content = replaces.get(key);
                exists =  exists || replaces.containsKey(key);
                if(null == content){
                    content = replaces.get(flag);
                    exists = exists || replaces.containsKey(flag);
                }
            }else{
                content = replaces.get(flag);
                exists =  exists || replaces.containsKey(flag);
            }
            if(null == content){
                content = "";
            }
            txt = txt.replace(flag, content);
        }
        return txt;
    }
}