# Office 文档服务设计

## 总体设计

Office 文档服务通过[Jacob (Java Com Bridge)](https://sourceforge.net/projects/jacob-project/) 调用Windows Office软件，提供文档操作和转换服务。
比如，替换文本、在指定书签处插入文本等，或者把WORD文档转换为PDF格式。

Office 文档服务的总体结构：

```
 +--------------------------------+
 | Windows Server OS              | 
 |                                |
 |     +----------------+         |
 |     | Office Service |         |
 |     +----------------+         |
 |             | call             |
 |     +-----------------+        |
 |     | Office Software |        |
 |     +-----------------+        |
 |                                |
 +--------------------------------+
```

## 服务场景

主要流程：

1. 客户端通过 API 接口向Office文档服务上传文档，并指明文档操作内容
2. Office文档服务通过Jacob调用Windows Office软件操作WORD文档
3. Office文档服务向客户端同步返回结果文档

## 接口设计

为减少交互，一个API 接口请求包括了文档和操作内容，通过编码类型`multipart/form-data`的POST请求完成。

例子：

```
POST /word-documents
Content-Type: multipart/form-data; boundary=--385716604555032261387884

----385716604555032261387884
Content-Disposition: form-data; name="processing"

{"actions": [{"method": "insertInBookmarks","args": { "sq16": "2020", "sq17": "5", "sq18": "25" }},{"method": "replaceText","args": { "oldText": "oldText", "newText": "newText" }}],"targetFileName": "new.pdf"}
----385716604555032261387884
Content-Disposition: form-data; name="file"; filename="1.doc"
Content-Type: application/msword

<文件内容>
```

{"actions": [{"method": "insertInBookmarks","args": { "sq16": "2020", "sq17": "10", "sq18": "cccc" }},{"method":"replaceText","args":{"oldText":"oldText","newText":"newText"}}],"targetFileName":"text.pdf"}

{"actions": [{"method": "addWaterMark","args": { "abscissa": "20", "ordinate": "20", "waterMarkCode": "t","waterMarkName":"1" }}],"targetFileName": "text1.docx"}

{"actions": [{"method": "addWaterMark","args": { "abscissa": "20", "ordinate": "20", "waterMarkCode": "t","waterMarkName":"1" }}],"targetFileName": "text1.docx"}

操作内容`processing`的请求体是一个序列化的JSON对象，定义例子如下：

```
  processing: {
    actions: [
      {
        // 方法：在指定书签处的内部插入文本
        method: "insertInBookmarks",
        args: { "sq16": "2020", "sq17": "5", "sq18": "25" }
      },
      {
        // 方法：替换文本
        method: "replaceText",
        args: { "oldText": "oldText", "newText": "newText" }
      }
    ],

    // 转换后的目标文件名，支持 .doc/.docx/.pdf 等3种格式. 必填
    "targetFileName": "new.pdf"
  }
```

参考: 
- [POSTMAN for Multipart/form-data](https://stackoverflow.com/questions/44182746/postman-for-multipart-form-data) May 25 '17
- [How to send the Multipart file and json data to spring boot](https://stackoverflow.com/questions/52818107/how-to-send-the-multipart-file-and-json-data-to-spring-boot) Oct 15 '18

## 处理逻辑

启动
  从启动参数/配置属性中解析临时目录
  从jar中复制jacob-xx.dll 到临时目录
  设置 jacob dll path 环境变量
  通过系统接口加载dll

OfficeController: Rest API

入口处理
  有 processing && processing.targetFileName, 转换后的目标文件必须是 .doc/.docx/.pdf 等3种格式
  上传文件须为 .doc/.docx 格式，存为临时目录下的临时文件

调用 Office 文档服务，操纵文档并保存为目标文件格式

向客户端返回目标文件，根据目标文件格式设置响应头 content-type 属性值


### Profile-specific Properties与日志输出

application-default.properties: 生产环境默认属性

application-dev.properties: 开发环境属性

调试时加程序参数：
```
--spring.profiles.active=dev

# 命令行例子
java -Dlogging.level.cn.innoway=DEBUG -jar office-service.jar --spring.profiles.active=dev
```



## TODO

### 线程池 (2020-6-2)

必须限制启动的Office程序数量，同时避免重复启动程序。故用线程池机制，已启动的程序始终在线程池内。

- v1：使用 semaphore 控制进入操纵WORD程序段的请求数量。但问题是每次进入后，都要启停WORD。

  BTW: 调试中发现WPS支持WORD自动化接口。机器上先装了WORD，再装了WPS，也能完成WORD文档操作服务

- v2: 使用 BlockingQueue 队列 (完成)
  初始化 BlockingQueue 队列，成员数量从配置属性中获取
  收到请求，进入预处理
  再从 BlockingQueue 队列中取可用服务
  操作WORD文档
  把服务放回队列中

### 启用 log4j2 异步日志 (Asynchronous Loggers)

  cf. [Asynchronous Loggers for Low-Latency Logging](https://logging.apache.org/log4j/2.x/manual/async.html)

  [LMAX Disruptor](https://lmax-exchange.github.io/disruptor/)  High Performance Inter-Thread Messaging Library


properties binding @Value
    Command Line Properties
    application.properties
    default value


## 踩坑小记

### 指定端口号启动spring boot jar程序 (2020-6-4)

```bash
java -jar <path/to/my/jar> --server.port=7788

# 或者，顺序不能错，-D... 必须置于 -jar 参数之前
java -Dserver.port=7788 -jar <path/to/my/jar>

# 或者，修改 application.properties
server.port=7788
```

### 设置日志输出级别

VM选项：

```text
-Dlogging.level.cn.innoway=DEBUG
```

cf. [Logging in Spring Boot](https://www.baeldung.com/spring-boot-logging) April 26, 2020

### 如何打包本地Jar和DLL库

Jacob库包括Jar和Windows DLL库文件，没有自身的POM文件。在项目POM中，通过`scope`和`systemPath`属性
把Jar加入到了项目依赖，而DLL是在应用启动时动态加载的。项目POM配置如下：

```text
  <dependency>
      <groupId>local</groupId>
      <artifactId>jacob</artifactId>
      <version>1.19</version>
      <scope>system</scope>
      <systemPath>
          <!-- Jar位于项目子目录 libs/jacob-1.19 中 -->
          ${project.basedir}/libs/jacob-1.19/jacob.jar
      </systemPath>
  </dependency>
```

但用maven打包时，生成的项目包中没有Jacob Jar和Windows DLL库文件。

#### 打包本地Jar

**方法一：maven-install-plugin**

按照这篇文章 [How to package spring-boot to jar with local jar using maven](https://stackoverflow.com/questions/48435305/how-to-package-spring-boot-to-jar-with-local-jar-using-maven) (Jan 25 '18) 加入plugin：

```text
    <plugin>
        <groupId>org.apache.maven.plugins</groupId>
        <artifactId>maven-install-plugin</artifactId>
        <version>2.5.2</version>
        <configuration>
            <groupId>local</groupId>
            <artifactId>jacob</artifactId>
            <version>1.19</version>
            <packaging>jar</packaging>
            <file>${project.basedir}/libs/jacob-1.19/jacob.jar</file>
        </configuration>
        <executions>
            <execution>
                <id>install-lib</id>
                <goals>
                    <goal>install-file</goal>
                </goals>
                <phase>validate</phase>
            </execution>
        </executions>
    </plugin>
```

然后clean一下，再注释 jacob 依赖的`scope`和`systemPath`属性。这样maven打包时`jacob.jar`可打包进去。

**方法二：项目maven库 (有问题)**

这个方法更简单：

```text
    <repositories>
        <repository>
            <id>local-repository</id>
            <url>file:///${project.basedir}/libs/jacob-1.19</url>
        </repository>
    </repositories>

    <dependencies>
        <!-- ... -->

        <dependency>
            <groupId>local</groupId>
            <artifactId>jacob</artifactId>
            <version>1.19</version>
        </dependency>
    </dependencies>
```

再运行:

```bash
mvn deploy:deploy-file -DgroupId=local -DartifactId=jacob -Dversion=1.19 -Durl=file:./libs/jacob-1.19/ -DrepositoryId=local-repository -DupdateReleaseInfo=true -Dfile=D:\workspace\innoway\project\NBSNEW\branches\DEV\office-service\libs\jacob-1.19\jacob.jar

# 或
mvn org.apache.maven.plugins:maven-install-plugin:2.5.2:install-file  ^
    -Dfile=D:\workspace\innoway\project\NBSNEW\branches\DEV\office-service\libs\jacob-1.19\jacob.jar ^
    -DgroupId=local -DartifactId=jacob -Dversion=1.19 ^
    -DlocalRepositoryPath=${master_project}/local-repository
```

**方法三：上传到内部Nexus Maven服务器**

指定POM依赖如下：

```text
  <dependency>
      <groupId>com.jacob</groupId>
      <artifactId>jacob</artifactId>
      <version>1.19</version>
  </dependency>
```

#### 打包DLL库

加入`resource`，指定要复制的文件：

```text
    <build>
        <resources>
            <resource>
                <directory>libs/jacob-1.19</directory>
                <includes>
                    <include>jacob*.dll</include>
                </includes>
            </resource>
        </resources>
        <!-- ... -->
    </build>
```

下面这种方法更简单一些：

把`jacob*.dll`移动到`src/main/resources/libs/jacob-1.19`目录。调试或打包时`resources`目录下的全部文件默认都会复制。

#### 载入DLL库

TODO: 使用 NativeUtils 载入DLL库，临时目录创建后，退出程序时无法删除


参考资料：

- [3 ways to add local jar to maven project](http://roufid.com/3-ways-to-add-local-jar-to-maven-project/)

  三种方法：
  * Adding the dependency as system scope
  * Creating a different local Maven repository
  * Using a Nexus repository manager

- [Class JarClassLoader](http://www.jdotsoft.com/JarClassLoader.php)

  The class loader to load classes, native libraries and resources from the top JAR and from JARs inside the top JAR.

- [native-utils](https://github.com/adamheinrich/native-utils)

  A simple library class which helps with loading dynamic JNI libraries stored in the JAR archive

- [JNI的替代者—使用JNA访问Java外部功能接口](https://www.cnblogs.com/lanxuezaipiao/p/3635556.html)

### 设置 docx 响应头的 content-type (2020-6-5)

通过Spring Boot编程配置添加 docx 的 mimetype:

```java
// DocxMimeMappingConfig.java
@Configuration
public class DocxMimeMappingConfig
    implements
      WebServerFactoryCustomizer<ConfigurableServletWebServerFactory> {
  @Override
  public void customize(final ConfigurableServletWebServerFactory factory) {
    final MimeMappings mappings = ((AbstractServletWebServerFactory) factory).getMimeMappings();
    mappings.add("docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    factory.setMimeMappings(mappings);
  }
}
```

再根据文件名取到 mimeType：

```java
  final String mimeType = request.getServletContext().getMimeType(filePath);
  response.setHeader("Content-Type", mimeType);
```

参考:
- [Spring Boot, static resources and mime type configuration](https://stackoverflow.com/questions/47222191/spring-boot-static-resources-and-mime-type-configuration)   Nov 13 '17
- [Getting a File’s Mime Type in Java](https://www.baeldung.com/java-file-mime-type) March 7, 2020


### 在程序启动时获取属性值 (2020-6-6)

使用 `@Value` 注解可在Spring管理的bean中注入属性值。例如：

```java
@RestController
public class OfficeController {
  @Value("${app.temp-path:abc}")
  private String tempPath;
  // ...
}
```

但是，如何在进入main时就取到属性值呢？

只有在spring context初始化后，才能以spring方式获取到属性值。否则，就只能手工读取属性文件。

获取方法有很多，我们选用CommandLineRunner接口方法。在这个接口的 run 方法中，spring context已初始化。

参考:
- [Spring Boot application to read a value from properties file in main method](https://stackoverflow.com/questions/48155833/spring-boot-application-to-read-a-value-from-properties-file-in-main-method) Jan 8 '18 at 18:12
- [How to assign a value from application.properties to a static variable?](https://stackoverflow.com/questions/45192373/how-to-assign-a-value-from-application-properties-to-a-static-variable)  Jul 19 '17
- [Spring Boot Reference Documentation - 4.1.7. Application Events and Listeners](https://docs.spring.io/autorepo/docs/spring-boot/current/reference/htmlsingle/#boot-features-application-events-and-listeners)

### 关于 Spring Boot Embedded Tomcat Logs (2020-6-7)

在 `application.properties` 开启Tomcat访问日志：
```
server.tomcat.accesslog.enabled=true
```

日志文件保存在用户临时目录中，例如：

```
C:\Users\<用户名>\AppData\Local\Temp\tomcat.6347928308270268076.8080\logs
```

参考：
- [Spring Boot Reference Documentation - Common Application properties - 11. Server properties](https://docs.spring.io/spring-boot/docs/current/reference/html/appendix-application-properties.html#server-properties)
- [Spring Boot Embedded Tomcat Logs](https://www.baeldung.com/spring-boot-embedded-tomcat-logs)
