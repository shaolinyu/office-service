
package cn.innoway.application;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;

import cn.innoway.msoffice.WordProcessing;

@SpringBootTest
class OfficeApplicationTests {

  @Test
  void contextLoads() {
  }
  
  public static void main(String[] args) {
    String process = "{\"actions\": [{\"method\": \"insertInBookmarks\",\"args\": { \"sq16\": \"2020\", \"sq17\": \"5\", \"sq18\": \"25\" }},{\"method\": \"repalceText\",\"args\": { \"oldText\": \"oldText\", \"newText\": \"newText\" }}],\"targetFileName\": \"转换结果.pdf\"}";
    ObjectMapper objectMapper = new ObjectMapper();
    try {
        WordProcessing wordProcessing = objectMapper.readValue(process, WordProcessing.class);
        System.out.println(wordProcessing.getTargetFileName());
    } catch (JsonMappingException e) {
        // TODO Auto-generated catch block
        e.printStackTrace();
    } catch (JsonProcessingException e) {
        // TODO Auto-generated catch block
        e.printStackTrace();
    }
  }
}
