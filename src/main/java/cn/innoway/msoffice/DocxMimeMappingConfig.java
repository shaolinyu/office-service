
package cn.innoway.msoffice;

import org.springframework.boot.web.server.MimeMappings;
import org.springframework.boot.web.server.WebServerFactoryCustomizer;
import org.springframework.boot.web.servlet.server.AbstractServletWebServerFactory;
import org.springframework.boot.web.servlet.server.ConfigurableServletWebServerFactory;
import org.springframework.context.annotation.Configuration;

/**
 * Spring Boot 2 容器配置：添加 docx mimetype
 * 
 */
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
