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


import org.anyline.office.docx.util.DocxUtil;
import org.dom4j.Element;

public class TagElement {
    private String text;
    private Element element;
    private int index; //element在top中可见内容中的下标 注意foot时区分是在首行(只有一行)还是尾行
    private boolean last; //在可见内容中是否第一个
    private boolean first;//在可见内容中是否最后个

    public String text() {
        return text;
    }

    public void text(String text) {
        this.text = text;
    }

    public int index() {
        return index;
    }

    public void index(int index) {
        this.index = index;
    }

    public boolean last() {
        return last;
    }

    public void last(boolean last) {
        this.last = last;
    }

    public boolean first() {
        return first;
    }

    public void first(boolean first) {
        this.first = first;
    }
    public Element element() {
        return element;
    }

    public void element(Element element) {
        this.element = element;
        if(null != element){
            text(element.getText());
        }
    }

    /**
     * 如果删除后上级为空后，是否删除上级
     * @param clear boolean
     */
    public void remove(boolean clear){
        if(clear) {
            DocxUtil.remove(element);
        }else{
            Element parent = element.getParent();
            if(null != parent){
                parent.remove(element);
            }
        }
    }
    public void remove(){
        remove(true);
    }
}
