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

import org.anyline.office.docx.entity.WDocument;
import org.anyline.office.docx.util.DocxUtil;
import org.anyline.office.util.TagUtil;
import org.dom4j.Element;

import java.util.ArrayList;
import java.util.List;

public class TagBox {
    private String name;
    private TagElement head;
    private TagElement foot;

    private List<Element> tops;
    private List<Element> templates = new ArrayList<>(); //tops中删除标签外内容
    private List<Element> contents;
    private WDocument doc;
    public TagBox(WDocument doc){
        this.doc = doc;
    }
    public TagElement head() {
        return head;
    }

    public void head(TagElement head) {
        this.head = head;
        if(null != head){
            String text = head.text();
            if(null != text){
                name(TagUtil.name(text, doc.namespace()+":"));
            }
        }
    }

    public TagElement foot() {
        return foot;
    }

    public void foot(TagElement foot) {
        this.foot = foot;
    }

    public List<Element> tops() {
        return tops;
    }

    public void tops(List<Element> tops) {
        this.tops = tops;
        int size = tops.size();
        for(int i=0; i<size; i++){
            Element top = tops.get(i);
            Element copy = top.createCopy();
            copy.setParent(top.getParent());
            //删除head前foot后内容
            List<Element> contents = DocxUtil.contents(copy);

            if (i == size - 1) {
                if (null != foot) {
                    for (int r = foot.index() + 1; r < contents.size(); r++) {
                        DocxUtil.remove(contents.get(r));
                    }
                }
            }
            if (i == 0) {
                for(int r = 0; r < head.index(); r++){
                    DocxUtil.remove(contents.get(r));
                }
            }
            templates.add(copy);
        }
    }

    public String name() {
        return name;
    }

    public void name(String name) {
        this.name = name;
    }


    /**
     * 如果删除后上级为空后，是否删除上级
     * @param clear boolean
     */
    public void remove(boolean clear){
        //不要删除tops top中可能有标签外的内容
        for(Element content:contents){
            if(clear){
                DocxUtil.remove(content);
            }else{
                Element parent = content.getParent();
                if(null != parent){
                    parent.remove(content);
                }
            }
        }
    }

    /**
     * 删除head foot
     */
    public void cut(){
        foot.remove();
        head.remove();
    }
    public void remove(){
        remove(true);
    }
    public List<Element> contents(){
        return contents;
    }
    public void contents(List<Element> contents){
        this.contents = contents;
    }
    public List<Element> templates(){
        return templates;
    }
}
