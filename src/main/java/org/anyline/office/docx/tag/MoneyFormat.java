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

import org.anyline.util.BasicUtil;
import org.anyline.util.MoneyUtil;

public class MoneyFormat extends AbstractTag implements Tag{
    public void release(){
        super.release();
    }
    @Override
    public String parse(String text) {
        String result = "";
        //<aol:money value="${total}"></aol:money>
        String key = fetchAttributeValue(text, "value", "v");
        if(BasicUtil.isEmpty(key)){
            return "";
        }
        Object data = context.data(key.trim());
        if(BasicUtil.isNotEmpty(data)) {
            double d = BasicUtil.parseDouble(data, 0d);
            result = MoneyUtil.format(d);
        }
        return result;
    }
}
