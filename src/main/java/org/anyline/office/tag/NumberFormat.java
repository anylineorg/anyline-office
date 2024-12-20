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

import org.anyline.util.BasicUtil;
import org.anyline.util.NumberUtil;

public class NumberFormat extends AbstractTag implements Tag{
    public void release(){
        super.release();
    }
    @Override
    public void run() {
        String result = null;
        //<aot:number format="###,##0.00" value="${total}"></aot:number>
        String format = fetchAttributeString("format", "f");
        Object data = fetchAttributeData("value");
        if(null == data){
            data = body(text, "number");
        }
        if(BasicUtil.isNotEmpty(data)){
            if (BasicUtil.isNotEmpty(format)) {
                result = NumberUtil.format(data.toString(), format);
            } else {
                result = data.toString();
            }
        }
        output(result);
    }
}
