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

import ognl.Ognl;
import ognl.OgnlContext;
import ognl.OgnlException;
import org.anyline.util.BasicUtil;
import org.anyline.util.DefaultOgnlMemberAccess;
import org.anyline.util.regular.RegularUtil;

public class If extends AbstractTag implements Tag {
    public String parse(String text) throws Exception{
        String html = "";
        String test = RegularUtil.fetchAttributeValue(text, "test");
        String value = RegularUtil.fetchAttributeValue(text, "value");
        if(BasicUtil.checkEl(test)){
            test = test.substring(2, test.length()-1);
        }
        String elseValue = RegularUtil.fetchAttributeValue(text, "else");
        if(null == elseValue){
            elseValue = "";
        }
        boolean chk = false;
        try {
            OgnlContext context = new OgnlContext(null, null, new DefaultOgnlMemberAccess(true));
            Boolean bol = (Boolean) Ognl.getValue(test, context, variables);
            if(null != bol){
                chk = bol;
            }
        } catch (OgnlException e) {
            e.printStackTrace();
        }
        if(chk){
            if(BasicUtil.isNotEmpty(value)){
                html = value;
            } else {
                //test中会有>影响表达式
                text = text.replace(test, "");
                html = RegularUtil.fetchTagBody(text, "aol:if");
            }
        }else{
            html = elseValue;
        }
        return html;
    }
}
