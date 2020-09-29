# 项目说明
本项目提供Word文档操作服务。服务程序在Windows环境下运行，需要安装Office 2010+或 Office 2007加相应插件。

## 如何使用

### 部署安装步骤
1. 编译打包，生成`office-service.jar`。
2. 编写运行批处理文件 `start.bat`，内容如下：
   ```bat
   TITLE Office Service (%cd%)
   java -Dlogging.level.cn.innoway=DEBUG -jar office-service.jar
   ```
3. 把上面两个文件压缩到`office-service.zip`文件中，复制到部署机器上，再解压到部署机器的`office-service`目录下。
4. 运行`start.bat`，启动服务。

### 使用Word文档操作服务
服务以Rest API接口方式提供，支持的操作内容包括：
- 在Word文档指定书签位置内插入文本
- 替换Word文档内的文本
- 生成 `.doc`、`.docx`或`.pdf`格式文档

先用工具测试，看看服务效果。

准备一个测试文档`test.doc`，其中含3个书签：sq16、sq17和sq18。另外去下载网络小工具`curl.exe`。然后，打开命令行窗口，运行命令：
```
curl --location --request POST "http://127.0.0.1:8080/watermark-documents" ^
--form "processing= {\"actions\": [{\"method\": \"addWaterMark\",\"args\": { \"abscissa\": \"20\", \"ordinate\": \"20\", \"waterMarkCode\": \"t\",\"waterMarkName\":\"1\" }}]}" ^
--form "file=@hello.doc" -o 转换结果.doc
```

```
curl --location --request POST "http://192.9.200.121:8080/word-documents" ^
--form "processing= {\"actions\": [{\"method\": \"insertInBookmarks\",\"args\": { \"sq16\": \"2020\", \"sq17\": \"5\", \"sq18\": \"25\" }},{\"method\": \"repalceText\",\"args\": { \"oldText\": \"oldText\", \"newText\": \"newText\" }}],\"targetFileName\": \"转换结果.pdf\"}" ^
--form "file=@test.doc" -o 转换结果.pdf
```

注，命令中的"^"符号为Windows命令的续行符。

最后，打开 “`转换结果.pdf`” 文件，看看在指定书签处是否有2020、5和25等3处文本。



服务以Rest API接口方式提供，支持的操作内容包括

```
curl --location --request POST "http://192.9.200.121:8080/watermark-documents" ^
--form "processing= {\"actions\": [{\"method\": \"repalceText\",\"args\": { \"oldText\": \"oldText\", \"newText\": \"newText\" }}],\"targetFileName\": \"转换结果.pdf\"}" ^
--form "file=@test.doc" -o 转换结果.pdf
```

注，命令中的"^"符号为Windows命令的续行符。

最后，打开 “`转换结果.pdf`” 文件，看看在指定书签处是否有2020、5和25等3处文本。

### 服务接口说明

Word文档操作服务接口采用 Rest API 方式，使用POST方法，请求体用`multipart/form-data`封装，其中包括了操作内容和文件内容。
请求体的操作内容块的标识名称 (name) 为`processing`，为JSON对象；
文件块的标识名称为`file`，源文件名由 `filename` 属性指定，源文件必须是Word文档格式，后缀名为 `.doc` 或 `.docx`。

操作内容块 `processing` 的JSON对象属性说明如下：

1、word接口 http://127.0.0.1:8080/word-documents
说明：仅支持图片水印，运行jar包后，将水印图片放入office_service_temp\docs\water.png
参数说明： file(上传文件) , processing
```
{
"actions":[
{
"method":"addWaterMark", //调用方法
"args":{
"abscissa":"100", //横坐标
"ordinate":"350" //纵坐标
}
}
],
"targetFileName":"new.doc"  //另存为文件名
}

```
2、excel接口说明 http://127.0.0.1:8080/excel-documents
```
{
"actions":[
{
"method":"addWaterMark",
"args":{
"waterContent":"水印名称", 添加水印的字
"size":"20", 大小
"left":"200", 横坐标
"right":"300" 纵坐标
}
}
],
"targetFileName":"new.xlsx"
}
```
3、ppt接口说明  http://127.0.0.1:8080/ppt-documents
````
{
"actions":[
{
"method":"addWaterMark",
"args":{
"waterContent":"水印名称",
"size":"30",
"left":"200",
"right":"300"
}
}
],
"targetFileName":"max111.ppt"
}
````

```
{
  actions: [
    {
      // 方法：在指定书签处的内部插入文本
      method: "insertInBookmarks",
      // 方法参数：书签名-文本值 对
      args: { "sq16": "2020", "sq17": "5", "sq18": "25" }
    },
    {
      // 方法：替换文本
      method: "replaceText",
      // 方法参数：oldText 旧文本，newText 新文本
      args: { "oldText": "oldText", "newText": "newText" }
    }
  ],
  
  // 转换后的目标文件名，支持 .doc/.docx/.pdf 等3种格式. 必填
  "targetFileName": "result.pdf"
}
```

请求例子：

```
POST /word-documents
Content-Type: multipart/form-data; boundary=--385716604555032261387884

----385716604555032261387884
Content-Disposition: form-data; name="processing"

{"actions": [{"method": "insertInBookmarks","args": { "sq16": "2020", "sq17": "5", "sq18": "25" }},{"method": "replaceText","args": { "oldText": "oldText", "newText": "newText" }}],"targetFileName": "result.pdf"}
----385716604555032261387884
Content-Disposition: form-data; name="file"; filename="test.doc"
Content-Type: application/msword

(文件内容略)
----385716604555032261387884--
```

响应例子：

```
HTTP/1.1 200
content-disposition: attachment;filename=result.pdf
Content-Type: application/pdf

(文件内容略)
```

## 运行和配置说明

### 运行环境和目录
Word文档操作服务程序在Windows环境下运行，预先安装Office 2010+。或者，安装 Office 2007，并加装插件`Save as PDF or XPS add-in`，这个插件在
Office 2007 SP2 包中也附带了。

运行Word文档操作服务程序 `office-service.jar`，默认会在当前目录下生成两个临时子目录：
- office_service_temp，用于存放jacob库文件，以及上传文件和转换结果文件
- tomcat，程序内嵌的Tomcat运行目录，其中有应用日志文件 `tombat/logs/office-service.log`

### 配置参数
程序的全部运行参数参见 `application-default.properties` 文件。

这里给出几个重要参数的说明：

- `server.port`，服务端口 ，默认为`8080`
- `app.temp-doc-path`，存放上传文件和转换结果文件的临时目录，默认为 `office_service_temp/docs`。
由于Word文档操作服务必须针对文件进行，而在物理硬盘上存取文件比较耗时。所以实际部署时可以使用内存模拟磁盘，把文件放在内存模拟盘上操作。
用这个参数可把文件临时目录指定为内存模拟盘上的目录。
- `app.show-word-window`，是否显示Word程序窗口，默认不显示。
- `app.word-instance-count`，Word程序并发数量，默认 5 个。每次文档服务请求都会启动一次Word程序，这个参数就是用于控制可同时运行的Word程序数量。
  **特别注意**，这个参数必须根据服务器性能而定。
- `logging.level.cn.innoway`，指定应用`cn.innoway`包的日志级别，这是Spring Boot的参数，建议设置为`DEBUG`。

命令例子：
```
java -Dapp.temp-doc-path=C:/temp/mydocs -Dapp.show-word-window=true -Dapp.word-instance-count=1 ^
-Dlogging.level.cn.innoway=DEBUG -jar office-service.jar --server.port=8899
```

命令例子的参数含义：
- 文件的临时目录为`C:\temp\mydocs`
- 显示打开的Word程序窗口
- Word程序并发数量为1
- 应用`cn.innoway`包的日志级别为DEBUG
- 服务端口为 `8899`


