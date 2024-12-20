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
        if(BasicUtil.isEmpty(data) || BasicUtil.isEmpty(var)){
            return;
        }

        if (null != data) {
            doc.variable(var, data);
            doc.replace(var, data.toString());
        }else{
            doc.replace(var, "");
            doc.variable(var, null);
        }
        DocxUtil.remove(contents);
    }
}
