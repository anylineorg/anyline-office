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

import org.anyline.entity.DataSet;
import org.anyline.util.BasicUtil;

public class Group extends AbstractTag implements Tag{
    private String by;

    public void release(){
        super.release();
        var = null;
        data = null;
        by = null;
    }
    public void run(){
        var = fetchAttributeString(head, "var");
        by = fetchAttributeString(head, "by");
        data = data();
        DataSet groups = null;
        if(BasicUtil.isNotEmpty(data) && BasicUtil.isNotEmpty(var) && BasicUtil.isNotEmpty(by)){
            if(data instanceof DataSet){
                DataSet set = (DataSet) data;
                groups = set.group(by.split(","));
            }
        }
        output(groups);
    }
}
