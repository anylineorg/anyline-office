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

package org.anyline.office;

import org.anyline.office.docx.entity.WDocument;
import org.anyline.util.DateUtil;
import org.anyline.util.FileUtil;

import java.io.File;

public class CheckBoxTest {
    public static void main(String[] args) throws Exception{
        WDocument doc = doc();
        doc.replace("FILE_URL_COL", "http://file.deepbit.cn/wechat/cefd43d9a92b2cc1e4bc9c0fbc3d51b1/20241118/IZ3SVe51rXRyMpX7.jpg");
        doc.replace("current_time", DateUtil.format());
        doc.save();
    }
    public void run(){

    }

    private static WDocument doc(){
        File file = new File("E:\\template\\tag.docx");
        File copy = new File("E:\\template\\result\\tag_"+System.currentTimeMillis()+".docx");
        FileUtil.copy(file, copy);
        WDocument doc = new WDocument(copy);
        System.out.println("create doc file:"+copy.getAbsolutePath());
        return doc;
    }


}
