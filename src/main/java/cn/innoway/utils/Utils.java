
package cn.innoway.utils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.FileSystemNotFoundException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.ProviderNotFoundException;
import java.nio.file.StandardCopyOption;

public class Utils {

  /**
   * The minimum length a prefix for a file has to have according to
   * {@link File#createTempFile(String, String)}}.
   */
  private static final int MIN_PREFIX_LENGTH = 3;

  /**
   * Temporary directory which will contain the DLLs.
   */
  private static File temporaryDir;

  /**
   * 检查目录是否存在或创建目录
   * 
   * @param path
   *          待创建的目录
   * @return true = 目录创建成功或目录已存在
   */
  public static boolean checkOrCreateDir(String path) {
    File file = new File(path);
    return file.exists() || file.mkdirs();
  }

  /**
   * Loads library from current JAR archive source:
   *
   * The file from JAR is copied into system temporary directory and then loaded. The temporary file
   * is deleted after exiting. Method uses String as filename because the pathname is "abstract",
   * not system-dependent.
   *
   * @param path
   *          The path of file inside JAR as absolute path (beginning with '/'), e.g.
   *          /package/File.ext
   * @param tempPath
   *          the path of file to be copied to
   * @return absolute path for the library copied to temporay dir
   * @throws IOException
   *           If temporary file creation or read/write operation fails
   * @throws IllegalArgumentException
   *           If source file (param path) does not exist
   * @throws IllegalArgumentException
   *           If the path is not absolute or if the filename is shorter than three characters
   *           (restriction of {@link File#createTempFile(String, String)}).
   * @throws FileNotFoundException
   *           If the file could not be found inside the JAR.
   * @see <a href=
   *      "http://adamheinrich.com/blog/2012/how-to-load-native-jni-library-from-jar">http://adamheinrich.com/blog/2012/how-to-load-native-jni-library-from-jar</a>
   * @see <a href=
   *      "https://github.com/adamheinrich/native-utils">https://github.com/adamheinrich/native-utils</a>
   */
  public static String loadLibraryFromJar(final String path, String tempPath) throws IOException {

    if (null == path || !path.startsWith("/")) {
      throw new IllegalArgumentException("The path has to be absolute (start with '/').");
    }

    // Obtain filename from path
    final String[] parts = path.split("/");
    final String filename = (parts.length > 1) ? parts[parts.length - 1] : null;

    // Check if the filename is okay
    if (filename == null || filename.length() < MIN_PREFIX_LENGTH) {
      throw new IllegalArgumentException("The filename has to be at least 3 characters long.");
    }

    // Prepare temporary file
    temporaryDir = new File(tempPath);
    if (!temporaryDir.exists()) {
      temporaryDir.mkdirs();
    }

    final File temp = new File(temporaryDir, filename);

    try (final InputStream is = Utils.class.getResourceAsStream(path)) {
      Files.copy(is, temp.toPath(), StandardCopyOption.REPLACE_EXISTING);
    } catch (final IOException e) {
      temp.delete();
      throw e;
    } catch (final NullPointerException e) {
      temp.delete();
      throw new FileNotFoundException("File " + path + " was not found inside JAR.");
    }

    try {
      String libPath = temp.getAbsolutePath();
      System.load(libPath);
      return libPath;
    } finally {
      if (isPosixCompliant()) {
        // Assume POSIX compliant file system, can be deleted after loading
        temp.delete();
      } else {
        // Assume non-POSIX, and don't delete until last file descriptor closed
        temp.deleteOnExit();
      }
    }
  }

  private static boolean isPosixCompliant() {
    try {
      return FileSystems.getDefault().supportedFileAttributeViews().contains("posix");
    } catch (final FileSystemNotFoundException | ProviderNotFoundException | SecurityException e) {
      return false;
    }
  }
}
