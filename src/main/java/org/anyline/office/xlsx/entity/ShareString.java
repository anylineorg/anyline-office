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

import org.dom4j.Element;

import java.util.ArrayList;
import java.util.List;

public class ShareString {
    private int index;
    private String text;
    private List<XRun> runs = new ArrayList<>();

    public ShareString(){}
    public ShareString(String text){
        this.text = text;
    }
    public ShareString(Element element, int index) {
        this.index = index;
        parse(element);
    }
    public void parse(Element element){
        List<Element> runs = element.elements("r");
        for(Element run:runs){
            XRun xr = new XRun();
            Element property = run.element("rPr");
            if(null != property){
                XProperty xp = new XProperty(property);
                xr.property(xp);
            }
            Element text = run.element("t");
            if(null != text){
                xr.text(text.getTextTrim());
            }
            this.runs.add(xr);
        }
        Element txt = element.element("t");
        if(null != txt){
            text(txt.getTextTrim());
        }
    }

    public ShareString text(String text) {
        this.text = text;
        return this;
    }

    public String text() {
        return text;
    }
    public String texts() {
        if(null == runs){
            return text;
        }else{
            StringBuilder builder = new StringBuilder();
            for(XRun run:runs){
                builder.append(run.text());
            }
            return builder.toString();
        }
    }

    public ShareString index(int index) {
        this.index = index;
        return this;
    }

    public int index() {
        return index;
    }
    public List<XRun> runs() {
        return runs;
    }

    public ShareString runs(List<XRun> runs) {
        this.runs = runs;
        return this;
    }
    public ShareString add(XRun run) {
        this.runs.add(run);
        return this;
    }

}
/*	<si>
		<t>${a.b_.bc.d}</t>
		<phoneticPr fontId="1" type="noConversion"/>
	</si>
	<si>
		<r>
			<t>${a+</t>
		</r>
		<r>
			<rPr>
				<sz val="11"/>
				<color rgb="FFFF0000"/>
				<rFont val="宋体"/>
				<family val="3"/>
				<charset val="134"/>
				<scheme val="minor"/>
			</rPr>
			<t>c</t>
		</r>
		<r>
			<rPr>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="宋体"/>
				<family val="2"/>
				<charset val="134"/>
				<scheme val="minor"/>
			</rPr>
			<t>}</t>
		</r>
		<phoneticPr fontId="1" type="noConversion"/>
	</si>
*/