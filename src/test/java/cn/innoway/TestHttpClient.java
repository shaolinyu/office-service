package cn.innoway;

import cn.hutool.http.HttpRequest;
import com.google.common.io.ByteStreams;
import com.google.common.io.Files;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.mime.HttpMultipartMode;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.charset.Charset;
import java.util.Map;

public class TestHttpClient {

    //private static final String upload_url = "http://127.0.0.1:8080/excel-documents";
    private static final String upload_url = "http://127.0.0.1:8080/word-documents";


    public static void main(String[] args) throws IOException {
    }


    private static void testHttp() throws IOException {
        cn.hutool.http.HttpResponse execute = HttpRequest.post(upload_url).
                form("file", new File("C:\\Users\\shao\\Desktop\\vpn.doc"),
                        "processing", "{\"actions\": [{\"method\": \"addWaterMark\"," +
                                "\"args\": {\"abscissa\": \"20\", \"ordinate\": \"20\"}}]," +
                                "\"targetFileName\": \"text1.doc\"}").execute();
        InputStream inputStream = execute.bodyStream();
        OutputStream out = new FileOutputStream(new File("D:\\123.doc"));
        ByteStreams.copy(inputStream, out);
        out.close();
        inputStream.close();
    }


    /**
     * 使用httpclint 发送文件
     *
     * @param file 上传的文件
     * @return 响应结果
     * @author: qingfeng
     * @date: 2019-05-27
     */
    public static String uploadFile(String url, MultipartFile file,
                                    String fileParamName,
                                    Map<String, String> headerParams,
                                    Map<String, String> otherParams) {
        CloseableHttpClient httpClient = HttpClients.createDefault();
        String result = "";
        try {
            String fileName = file.getOriginalFilename();
            HttpPost httpPost = new HttpPost(url);
            //添加header
            for (Map.Entry<String, String> e : headerParams.entrySet()) {
                httpPost.addHeader(e.getKey(), e.getValue());
            }
            MultipartEntityBuilder builder = MultipartEntityBuilder.create();
            builder.setCharset(Charset.forName("utf-8"));
            builder.setMode(HttpMultipartMode.BROWSER_COMPATIBLE);//加上此行代码解决返回中文乱码问题
            builder.addBinaryBody(fileParamName, file.getInputStream(), ContentType.MULTIPART_FORM_DATA, fileName);// 文件流
            for (Map.Entry<String, String> e : otherParams.entrySet()) {
                builder.addTextBody(e.getKey(), e.getValue());// 类似浏览器表单提交，对应input的name和value
            }
            HttpEntity entity = builder.build();
            httpPost.setEntity(entity);
            HttpResponse response = httpClient.execute(httpPost);// 执行提交
            HttpEntity responseEntity = response.getEntity();
            if (responseEntity != null) {
                // 将响应内容转换为字符串
                result = EntityUtils.toString(responseEntity, Charset.forName("UTF-8"));
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                httpClient.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return result;
    }

}
