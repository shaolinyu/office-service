package cn.innoway.utils;

import cn.hutool.http.HttpUtil;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.StatusLine;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.mime.HttpMultipartMode;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.entity.mime.content.FileBody;
import org.apache.http.entity.mime.content.StringBody;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;

public class HttpUtils {

    private static final Logger LOGGER = LoggerFactory.getLogger(HttpUtils.class);

    private static final String url = "http://127.0.0.1:8080/word-documents";

    private static final String body = "{\"actions\": [{\"method\": \"addWaterMark\",\"args\": { \"abscissa\": \"100\", \"ordinate\": \"350\", \"waterMarkCode\": \"t\",\"waterMarkName\":\"1\" }}]," +
            "\"targetFileName\": \"max.doc\"}";



    public static void postRun() {
        Map<String, Object> map = new HashMap<>();
        map.put("processing", body);
        String post = HttpUtil.post(url, map, 10000);
        System.out.println(post);

    }

    public static void main(String[] args) throws IOException {
        CloseableHttpClient httpclient = HttpClients.createDefault();
        try {
            HttpPost httppost = new HttpPost(url);
            File file = new File("C:\\Users\\shao\\Desktop\\test.xlsx");
            FileBody bin = new FileBody(file);
            StringBody comment = new StringBody(body, ContentType.TEXT_PLAIN);

            HttpEntity reqEntity = MultipartEntityBuilder.create()
                    .addPart("file", bin)
                    .addPart("processing", comment)
                    .build();
            httppost.setEntity(reqEntity);
            System.out.println("executing request " + httppost.getRequestLine());
            CloseableHttpResponse response = httpclient.execute(httppost);
            try {
                System.out.println("----------------------------------------");
                System.out.println(response.getStatusLine());
                HttpEntity resEntity = response.getEntity();
                System.out.println(" resEntity.getContentType():"+ resEntity.getContentType());

                InputStream ins = resEntity.getContent();

                BufferedInputStream bi = new BufferedInputStream(ins);
                FileOutputStream fos = new FileOutputStream("D:\\max.doc");
                byte[] by = new byte[1024];
                int len = 0;
                while((len=bi.read(by))!=-1){
                    fos.write(by,0,len);
                }
                fos.close();
                bi.close();

                if (resEntity != null) {
                    System.out.println("Response content length: " + resEntity.getContentLength());
                }
                EntityUtils.consume(resEntity);
            } finally {
                response.close();
            }
        } finally {
            httpclient.close();
        }
    }

}
