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

import org.anyline.handler.Downloader;
import org.anyline.util.BasicUtil;
import org.anyline.util.regular.RegularUtil;

public class Include extends AbstractTag implements Tag{
	private static final long serialVersionUID = 1L;
	private String url;

	public void release(){
		super.release();
	}
	public void run() {
		try {
			String url = fetchAttributeString("url");
			String body = RegularUtil.fetchTagBody(text, doc.namespace()+":include");
			if(BasicUtil.isEmpty(url)){
				return;
			}
			url = url.trim();
			if(url.startsWith("http")){
				String host = doc.aoiHost();
				if(null != host){
					if(host.endsWith("/") && url.startsWith("/")){
						url = host.substring(1) + url;
					}else {
						if (!host.endsWith("/") && !url.startsWith("/")) {
							url = host + "/" + url;
						} else {
							url = host + url;
						}
					}
				}
			}
			Downloader downloader = doc.getDownloader();
			if(null != downloader){
				String content = downloader.post(url, body);
				if(null != content){
					content = content.trim();
					if(content.startsWith("<w:")){
						//open xml源生标签
					}else if(content.startsWith("<")){
						//html
						doc.parseHtml(tops.get(0), contents.get(0), content);
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
