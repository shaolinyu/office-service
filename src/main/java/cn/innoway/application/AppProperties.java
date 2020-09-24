
package cn.innoway.application;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

@Configuration
@ConfigurationProperties(prefix = "app")
public class AppProperties {

  //图片绝对地址
  private String picturePath = "office_service_temp/docs/1.png";

  // 临时目录默认为当前运行目录下的子目录 office_service_temp
  // lib子目录
  private String tempLibPath = "office_service_temp/libs/jacob/";

  // 文档子目录
  private String tempDocPath = "office_service_temp/docs";

  // 显示窗口
  private Boolean showWordWindow = false;

  // WORD实例数量
  private Integer wordInstanceCount = 5;

  public String getTempLibPath() {
    return tempLibPath;
  }

  public void setTempLibPath(final String tempLibPath) {
    this.tempLibPath = tempLibPath;
  }

  public String getTempDocPath() {
    return tempDocPath;
  }

  public void setTempDocPath(final String tempDocPath) {
    this.tempDocPath = tempDocPath;
  }

  public Integer getWordInstanceCount() {
    return wordInstanceCount;
  }

  public void setWordInstanceCount(final Integer wordInstanceCount) {
    if (wordInstanceCount < 1 || wordInstanceCount > 10) {
      throw new IllegalArgumentException(
          "Error value for property of wordInstanceCount: " + wordInstanceCount);
    }

    this.wordInstanceCount = wordInstanceCount;
  }

  public Boolean getShowWordWindow() {
    return showWordWindow;
  }

  public void setShowWordWindow(final Boolean showWordWindow) {
    this.showWordWindow = showWordWindow;
  }

  public String getPicturePath(){
      return  this.picturePath;
  }

}
