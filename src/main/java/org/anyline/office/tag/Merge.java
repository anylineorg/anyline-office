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

import org.anyline.office.docx.entity.WTc;
import org.anyline.office.docx.util.DocxUtil;
import org.dom4j.Element;


public class Merge extends AbstractTag implements Tag{
    private String by;//依赖的其他列,其他列的合并时当前列才合并
    private String ignores;//忽略合并的值 有可能所有空值已被换成了 特定符号
    private String scope = "tc"; //合并范围、tr:行内检测 行内多列值相同时合并列  tc:列内检测
    public void release(){
        super.release();
    }
    public void run(){
        scope = fetchAttributeString("scope", "s");
        by = fetchAttributeString("by");
        ignores = fetchAttributeString("ignores");
        if("tc".equalsIgnoreCase(scope) || "td".equalsIgnoreCase(scope)){
            //合并行 当前列范围内合并
            mergeRows();
        }else if("tr".equalsIgnoreCase(scope)){
            mergeCols();
        }
    }
    private void mergeRows(){
        Element table = DocxUtil.getParent(tops.get(0), "tbl");
        Element tc = DocxUtil.getParent(tops.get(0), "tc");
        WTc wtc = WTc.tc(tc);
        if(null == wtc){
            doc.tables(true);
            wtc = WTc.tc(tc);
        }
        int index = wtc.getRowspan();
    }
    private void mergeCols(){

    }

}
