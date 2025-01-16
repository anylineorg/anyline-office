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

import java.util.ArrayList;
import java.util.List;

public class Split extends AbstractTag implements Tag{
    public void release(){
        super.release();
    }
    @Override
    public void run() {
        String var = fetchAttributeString("var");
        String str = fetchAttributeString("value");
        if(BasicUtil.isEmpty(str)){
            str = body(text, "date");
        }
        String regex = fetchAttributeString("regex");
        if(BasicUtil.isEmpty(regex)){
            regex = ",";
        }
        if(BasicUtil.isNotEmpty(var) && BasicUtil.isNotEmpty(str)){
            List<String> list = new ArrayList<>();
            String[] tmps = str.split(regex);
            if(null != tmps) {
                for (String tmp : tmps) {
                    list.add(tmp);
                }
            }
            output(list);
        }
        box.remove();
    }
}
