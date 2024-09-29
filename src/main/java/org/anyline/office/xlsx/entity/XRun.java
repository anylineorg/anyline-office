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

public class XRun {
    private String text;
    private XProperty property;

    public String text() {
        return text;
    }

    public XRun text(String text) {
        this.text = text;
        return this;
    }

    public XProperty property() {
        return property;
    }

    public XRun property(XProperty property) {
        this.property = property;
        return this;
    }
}
/*
	<si>
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
* */