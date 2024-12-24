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
import org.anyline.util.BasicUtil;

public class Set extends AbstractTag implements Tag{
    public void release(){
        super.release();
        var = null;
        data = null;
    }
    public void run(){
        data = data();
        if(null == data){
            data = fetchAttributeData("value", "v");
        }
        if(BasicUtil.isBoolean(data)){
            data = BasicUtil.parseBoolean(data, false);
        }
        if(BasicUtil.isNumber(data)){
            data = BasicUtil.parseDecimal(data, null);
        }
        if(BasicUtil.isEmpty(data) || BasicUtil.isEmpty(var)){
            return;
        }
        String str = data .toString();
        if(str.startsWith("'") && str.startsWith("'")){
            data = str.substring(1, str.length()-1);
        }
        output(data);
        DocxUtil.remove(contents);
    }
}
