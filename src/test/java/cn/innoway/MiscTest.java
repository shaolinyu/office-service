
package cn.innoway;

import static org.junit.jupiter.api.Assertions.assertTrue;

import com.google.common.io.Files;
import java.util.Optional;
import org.junit.jupiter.api.Test;

public class MiscTest {

  @Test
  public void testExtension() {
    String name = null;
    String targetFileName = Optional.ofNullable(name).orElse("");
    String extension = Files.getFileExtension(targetFileName);
    assertTrue(extension.isEmpty());

    name = "abc";
    targetFileName = Optional.ofNullable(name).orElse("");
    extension = Files.getFileExtension(targetFileName);
    assertTrue(extension.isEmpty());

    name = "abc.";
    targetFileName = Optional.ofNullable(name).orElse("");
    extension = Files.getFileExtension(targetFileName);
    assertTrue(extension.isEmpty());

    name = "abc.docx";
    targetFileName = Optional.ofNullable(name).orElse("");
    extension = Files.getFileExtension(targetFileName);
    assertTrue(extension.equals("docx"));

    name = "def.pdf";
    targetFileName = Optional.ofNullable(name).orElse("");
    extension = Files.getFileExtension(targetFileName);
    assertTrue(extension.equals("pdf"));
  }

}
