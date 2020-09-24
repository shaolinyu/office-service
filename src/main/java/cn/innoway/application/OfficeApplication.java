
package cn.innoway.application;

import cn.innoway.utils.Utils;
import com.jacob.com.LibraryLoader;
import java.io.IOException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication(scanBasePackages = { "cn.innoway.application", "cn.innoway.msoffice" })
public class OfficeApplication implements CommandLineRunner {
  @Autowired
  private AppProperties appProperties;

  /**
   * 程序入口
   *
   * @param args
   *          参数
   */
  public static void main(final String[] args) {
    SpringApplication.run(OfficeApplication.class, args);
  }

  @Override
  public void run(final String... args) throws IOException {
    final String tempLibPath = appProperties.getTempLibPath();
    final String tempDocPath = appProperties.getTempDocPath();

    // 创建临时目录
    if (!Utils.checkOrCreateDir(tempLibPath) || !Utils.checkOrCreateDir(tempDocPath)) {
      throw new RuntimeException("Error creating temporay directories");
    }

    String arch = System.getProperty("sun.arch.data.model");
    System.out.println(arch);
    
    // 取当前 jacob 版本的DLL库名称，避免版本变动造成代码变动
    final String libFile = "/libs/jacob/" + LibraryLoader.getPreferredDLLName() + ".dll";
    System.out.println(libFile);
    // 从 jar 中复制 DLL 到临时目录，并载入之
    final String libPath = Utils.loadLibraryFromJar(libFile, tempLibPath);
    System.setProperty(LibraryLoader.JACOB_DLL_PATH, libPath);
  }
}
