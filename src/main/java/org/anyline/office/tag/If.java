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

import ognl.Ognl;
import ognl.OgnlContext;
import ognl.OgnlException;
import org.anyline.office.docx.entity.WTable;
import org.anyline.office.docx.entity.WTr;
import org.anyline.office.docx.util.DocxUtil;
import org.anyline.office.util.TagUtil;
import org.anyline.util.BasicUtil;
import org.anyline.util.DefaultOgnlMemberAccess;
import org.anyline.util.regular.RegularUtil;
import org.dom4j.Element;

/**
 * 注意以value为主 没有value的再读body
 */
public class If extends AbstractTag implements Tag {
    public void release(){
        super.release();
    }
    public void run() throws Exception{
        String test = RegularUtil.fetchAttributeValue(box.head().text(), "test");
        String value = fetchAttributeString("value", "v");
        String var = fetchAttributeString("var");
        //false时是否删除
        boolean remove = BasicUtil.parseBoolean(fetchAttributeString("remove", "rm"), false);
        //删除的对象 tc/td 或 tr
        String scope = fetchAttributeString("scope");
        if(BasicUtil.checkEl(test)){
            test = test.substring(2, test.length()-1);
        }
        String elseValue = fetchAttributeString("else");
        if(null == elseValue){
            elseValue = "";
        }
        boolean chk = false;
        try {
            OgnlContext ognl = new OgnlContext(null, null, new DefaultOgnlMemberAccess(true));
            Boolean bol = (Boolean) Ognl.getValue(test, ognl, context.variables());
            if(null != bol){
                chk = bol;
            }
        } catch (OgnlException e) {
            log.error("ognl表达式异常:{}", test);
            e.printStackTrace();
        }


        if(BasicUtil.isEmpty(var)) {
            if (chk) {
                if (BasicUtil.isNotEmpty(value)) {
                    //如果有value值
                    output(value);
                } else {
                    TagUtil.parse(doc, box.tops(), context);
                }
            } else {
                //html = elseValue;
                //删除body中的tops
                //TagUtil.clear(doc, tops);
                box.remove();
                if (remove) {//删除行或行
                    if ("tc".equalsIgnoreCase(scope) || "td".equalsIgnoreCase(scope)) {
                        Element tc = DocxUtil.getParent(contents.get(0), "tc");
                        Element tr = tc.getParent();
                        WTr wtr = WTr.tr(tr);
                        wtr.remove(tc);
                    } else if ("tr".equalsIgnoreCase(scope)) {
                        Element tr = DocxUtil.getParent(contents.get(0), "tr");
                        Element table = tr.getParent();
                        WTable wtable = WTable.table(table);
                        wtable.remove(tr);
                    }
                }
            }
        }else{
            //TagUtil.clear(doc, tops);
            box.remove();
            doc.variable(var, chk);
        }
    }

}
