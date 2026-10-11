PowerPoint Narrator — iPad Safari 网页版（开发测试版）

包含：
- 响应式 iPad 网页（大字体）
- 上传 PPTX
- 提取每页 Speaker Notes 为 TXT
- 检查和保留 PPTX 内已有音频（原文件直接返回以避免破坏媒体关系）
- 通过 OpenAI API 为每页 Notes 生成 MP3 并嵌入 PPTX 媒体对象

重要限制：
1. 本压缩包是网站源码，不是已经上线的网站。仅使用 iPad 的用户需要将源码部署到云端，才能通过 Safari 使用。
2. 目前没有 SayTXT 批量 API 集成；不能保证复制 SayTXT 专有声音。
3. AI 语音模式生成的媒体对象，在某些 iPad PowerPoint 版本中可能需要手动点击，尚未验证自动播放与自动换页。已有音频模式保留原文件的播放设置。
4. OpenAI API 使用独立计费密钥，不是 ChatGPT Plus 订阅权益。
5. 每段讲稿目前最多使用前 4096 个字符。处理 120 张会花时间并产生 API 成本。
6. 处理包含保密材料时，不建议部署为没有身份认证的公开服务。

云端部署（适用于仅用 iPad 的用户）：
- 在 GitHub 网页新建私人仓库，上传 app.py、index.html、requirements.txt、Dockerfile。
- 在支持 Docker 的云平台创建 Web Service，连接仓库并部署。
- 平台必须允许至少 150 MB 请求体和足够的处理时间；免费方案可能超时。
- 部署成功后在 iPad Safari 打开平台提供的 HTTPS 网址。
- 优先测试“只提取 Speaker Notes”，再测试“保留已有音频”。
- AI 模式涉及 API Key，建议在服务端配置认证和密钥代理后再使用。

本项目尚未进行云端部署；没有现成公网访问链接。
