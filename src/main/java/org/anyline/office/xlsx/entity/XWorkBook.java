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

import org.anyline.log.Log;
import org.anyline.log.LogProxy;
import org.anyline.util.DomUtil;
import org.anyline.util.ZipUtil;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;

import java.io.File;
import java.nio.charset.Charset;
import java.util.*;

public class XWorkBook {

    private static Log log = LogProxy.get(XWorkBook.class);
    private File file;
    private String charset = "UTF-8";
    private String xml = null;      // workbook.xml文本
    private org.dom4j.Document doc = null;
    private LinkedHashMap<String, String> replaces = new LinkedHashMap<>();
    /**
     * 文本原样替换，不解析原文没有${}的也不要添加
     */
    private LinkedHashMap<String, String> txt_replaces = new LinkedHashMap<>();
    private boolean autoMergePlaceholder = true;
    private List<ShareString> shares = new ArrayList<>();
    private LinkedHashMap<String, ShareString> shares_map = new LinkedHashMap<>();
    private org.dom4j.Document sharedDoc = null;
    LinkedHashMap<String, XSheet> sheets = new LinkedHashMap<>();

    public XWorkBook(File file){
        this.file = file;
    }

    public XWorkBook(String file){
        this.file = new File(file);
    }

    private void load(){
        if(null == xml){
            reload();
        }
    }
    public XSheet sheet(String key){
        return sheets.get(key);
    }
    public XSheet sheet(int index){
        if(sheets.isEmpty()){
            load();
        }
        int i = 0;
        for(XSheet sheet:sheets.values()){
            if(i == index){
                return sheet;
            }
            i ++;
        }
        return null;
    }
    public XSheet sheet(){
        return sheet(0);
    }

    public LinkedHashMap<String, XSheet> sheets(){
        return sheets;
    }

    public void reload(){
        try {
            xml = ZipUtil.read(file, "xl/workbook.xml", charset);
            doc = DocumentHelper.parseText(xml);
            List<String> items = ZipUtil.getEntriesNames(file);
            String shares = ZipUtil.read(file, "xl/sharedStrings.xml", charset);
            shares(shares);
            for(String item:items){
                if(item.contains("xl/worksheets") && item.endsWith(".xml")){
                    String name = item.replace("xl/worksheets/", "").replace(".xml", "");
                    Document doc = DocumentHelper.parseText(ZipUtil.read(file, item, charset));
                    XSheet sheet = new XSheet(this, doc, name);
                    sheets.put(name, sheet);
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }


    /**
     * 解析标签
     */
    public void parseTag(){
        for(XSheet sheet:sheets.values()){
            sheet.parseTag();
        }
    }
    /**
     * 设置占位符替换值 在调用save时执行替换<br/>
     * 注意如果不解析的话 不会添加自动${}符号 按原文替换,是替换整个文件的纯文件，包括标签名在内
     * @param parse 是否解析标签 true:解析HTML标签 false:直接替换文本
     * @param key 占位符
     * @param content 替换值
     */
    public void replace(boolean parse, String key, String content){
        if(null == key && key.trim().length()==0){
            return;
        }
        if(parse) {
            replaces.put(key, content);
        }else{
            txt_replaces.put(key, content);
        }
    }
    public void replace(String key, String content){
        replace(true, key, content);
    }

    public void save(){
        save(Charset.forName("UTF-8"));
    }
    public void save(Charset charset){
        try {
            //加载文件
            load();
            if(autoMergePlaceholder){
                mergePlaceholder();
            }
            Map<String, String> zip_replaces = new HashMap<>();
            for(XSheet sheet:sheets.values()){
                sheet.replace(true, replaces);
                sheet.replace(false, txt_replaces);
                zip_replaces.put("xl/worksheets/"+sheet.name()+".xml", DomUtil.format(sheet.doc()));
            }
            zip_replaces.put("xl/sharedStrings.xml", DomUtil.format(sharedDoc));
            ZipUtil.replace(file, zip_replaces, charset);
        }catch (Exception e){
            e.printStackTrace();
        }
    }
    //直接替换文本不解析
    public String replace(String text, Map<String, String> replaces){
        if(null != text){
            for(String key:replaces.keySet()){
                String value = replaces.get(key);
                //原文没有${}的也不要添加
                text = text.replace(key, value);
            }
        }
        return text;
    }

    /**
     * 合并点位符 ${key} 拆分到3个t中的情况
     * 调用完replace后再调用当前方法，因为需要用到replace里提供的占位符列表
     */
    public void mergePlaceholder(){
        List<String> placeholders = new ArrayList<>();
        placeholders.addAll(replaces.keySet());
        mergePlaceholder(placeholders);
    }
    /**
     * 合并点位符 ${key} 拆分到3个t中的情况
     * @param placeholders 占位符列表 带不还${}都可以 最终会处理掉${}
     */
    public void mergePlaceholder(List<String> placeholders){
    }
    public void mergePlaceholder(Element box, List<String> placeholders){
    }

    public void shares(String xml) throws Exception{
       sharedDoc = DocumentHelper.parseText(xml);
       Element root = sharedDoc.getRootElement();
       List<Element> list = root.elements();
       int index = 0;
       for(Element item:list){
           ShareString share = new ShareString(item, index ++);
           shares.add(share);
           String text = share.text();
           if(null != text) {
               shares_map.put(text, share);
           }
       }
    }
    public List<ShareString> shares(){
        return this.shares;
    }
    public ShareString share(int index){
        ShareString share = null;
        if(index>=0 && index < shares.size()){
            share = shares.get(index);
        }
        return share;
    }
    public int share(String text){
        ShareString share = shares_map.get(text);
        if(null != share){
            return share.index();
        }
        Element root = sharedDoc.getRootElement();
        Element si = root.addElement("si");
        Element t = si.addElement("t");
        t.setText(text);
        int index = root.elements("si").indexOf(si);
        share = new ShareString(si,  index);
        shares.add(share);
        shares_map.put(text, share);
        int size = shares.size();
        root.attribute("count").setValue(size+"");
        root.attribute("uniqueCount").setValue(size+"");
        return index;
    }

}
