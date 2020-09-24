
package cn.innoway.msoffice;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.util.Map;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.convert.converter.Converter;
import org.springframework.stereotype.Component;

/**
 * WORD文档处理类
 * 
 */
public class WordProcessing {
  private static final Logger logger = LoggerFactory.getLogger(WordProcessing.class);

  private Action[] actions;

  // 处理后的目标文件名
  private String targetFileName;

  public Action[] getActions() {
    return actions;
  }

  public void setActions(final Action[] actions) {
    this.actions = actions;
  }

  public String getTargetFileName() {
    return targetFileName;
  }

  public void setTargetFileName(final String targetFileName) {
    this.targetFileName = targetFileName;
  }

  public static class Action {

    // 操作方法名
    private String method;

    // 操作方法的参数
    private Map<String, String> args;

    public String getMethod() {
      return method;
    }

    public void setMethod(final String method) {
      this.method = method;
    }

    public Map<String, String> getArgs() {
      return args;
    }

    public void setArgs(final Map<String, String> args) {
      this.args = args;
    }
  }

  /**
   * 把请求体JSON对象 wordProcessing 转换为 WordProcessing 对象实例
   *
   */
  @Component
  public static class JsonToWordProcessingConverter implements Converter<String, WordProcessing> {

    @Autowired
    private ObjectMapper objectMapper;

    @Override
    public WordProcessing convert(final String source) {
      try {
        final WordProcessing wordProcessing = objectMapper.readValue(source, WordProcessing.class);
        return wordProcessing;
      } catch (final JsonProcessingException e) {
        logger.info("Oops! " + e.getMessage());
      }

      return null;
    }
  }
}
