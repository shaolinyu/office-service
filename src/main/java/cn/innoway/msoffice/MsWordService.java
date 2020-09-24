
package cn.innoway.msoffice;

import cn.innoway.application.AppProperties;
import com.google.common.collect.ImmutableMap;
import com.google.common.io.Files;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;

/**
 * Word文档服务类 修改自 jacob samples
 * <p>
 * Submitted to the Jacob SourceForge web site as a sample 3/2005
 * <p>
 * This sample is BROKEN because it doesn't call quit!
 *
 * @author Date Created Description Jason Twist 04 Mar 2005 Code opens a locally stored Word
 *         document and extracts the Built In properties and Custom properties from it. This code
 *         just gives an intro to JACOB and there are sections that could be enhanced
 */
public class MsWordService{
  // Declare word object
  private ActiveXComponent objWord;

  private Dispatch documents;

  // Current Active Document
  private Dispatch document;

  private Dispatch wordObject;

  /**
   * Empty Constructor
   */
  public MsWordService() {
  }

  /**
   * 文件格式是否支持
   *
   * @param fileName
   *          文件名
   * @return true = 支持，false = 不支持
   */
  public static boolean isSourceFileFormatSupported(final String fileName) {
    final String extension = Files.getFileExtension(fileName);
    return ("doc".equals(extension) || "docx".equals(extension)
            || "xls".equals(extension) || "xlsx".equals(extension)
            || "ppt".equals(extension) || "pptx".equals(extension));
  }

  /**
   * 启动Word
   *
   * @param showWindow
   *          if di splay window of Word
   */
  public void startWord(final boolean showWindow) {
    /*if (objWord != null ){
      return;
    }*/

    // Instantiate objWord
    objWord = new ActiveXComponent("Word.Application");

    // Assign a local word object
    wordObject = objWord.getObject();

    // Create a Dispatch Parameter to hide the document that is opened
    Dispatch.put(wordObject, "Visible", new Variant(showWindow));

    // Instantiate the Documents Proper
      // ty
    documents = objWord.getProperty("Documents").toDispatch();
  }

  /**
   * 打开文档
   *
   * @param filename
   *          file to be opened
   */
  public void open(final String filename) {
    if (documents == null) {
      throw new RuntimeException("Oops! Word not started");
    }

    // Open a word document, Current Active Document
    document = Dispatch.call(documents, "Open", filename).toDispatch();
  }

  /**
   * Closes a document
   *
   * @param save
   *          关闭时是否保存文档
   */
  public void close(final boolean save) {
    Dispatch.call(document, "Close", new Variant(save));
  }

  /**
   * 保存并关闭文档
   */
  public void saveAndClose() {
    close(true);
  }

  /**
   * 另存并关闭文档
   *
   * @param newFileName
   *          新文件名
   */
  public void saveAsAndClose(final String newFileName, final MsWordFormat msWordFormat) {
    Dispatch.call(document, "SaveAs", newFileName, msWordFormat.getValue());

    // 另存时，源文件不保存
      close(false);
  }

  /**
   * Close word app
   */
  public void quit() {
    if (objWord != null) {
      objWord.invoke("Quit", 0);
      objWord = null;
    }
  }

  private static final BiConsumer<MsWordService, Map<String, String>> METHOD_INSERT_IN_BOOKMARKS = MsWordService::insertInBookmarks;

  private static final BiConsumer<MsWordService, Map<String, String>> METHOD_REPLACE_TEXT = MsWordService::replaceText;

  private static final BiConsumer<MsWordService, Map<String, String>> METHOD_ADD_WATERMARK = MsWordService::addWaterMark;

  private static final Map<String, BiConsumer<MsWordService, Map<String, String>>> methods = ImmutableMap
      .of("insertInBookmarks", METHOD_INSERT_IN_BOOKMARKS, "replaceText", METHOD_REPLACE_TEXT,
          "addWaterMark", METHOD_ADD_WATERMARK);

  /**
   * 查询Word操作方法
   *
   * @param method
   *          方法名
   * @return 函数接口 (functional interface)
   */
  public static BiConsumer<MsWordService, Map<String, String>> getMethod(final String method) {
    return methods.get(method);
  }

  private static final List<String> methodList = Arrays.asList("insertInBookmarks", "replaceText",
      "addWaterMark");

  /**
   * 文档操作方法是否有效
   *
   * @param method
   *          方法名
   * @return true = 有效
   */
  public boolean isMethodValid(final String method) {
    return methodList.contains(method);
  }

  /**
   * 替换文本
   *
   * @param map
   *          替换的前后文本
   */
  public void replaceText(final Map<String, String> map) {
    final String oldText = map.get("oldText");
    final String newText = map.get("newText");
    if (oldText == null || oldText.isEmpty()) {
      return;
    }

    final Dispatch selection = objWord.getProperty("Selection").toDispatch();
    final Dispatch find = Dispatch.call(selection, "Find").toDispatch();
    Dispatch.put(find, "Text", oldText);
    Dispatch.call(find, "Execute");
    Dispatch.put(selection, "Text", newText);
  }

  protected Dispatch getBookmarksDispatch() {
    final Dispatch bookmarks = Dispatch.call(document, "Bookmarks").toDispatch();
    return bookmarks;
  }

  /**
   * 在指定书签处插入文本
   *
   * @param bookmarksText
   *          待插入的书签文本列表
   */
  public void insertAtBookmarks(final List<String[]> bookmarksText) {
    final Dispatch bookmarksDispatch = getBookmarksDispatch();
    for (final String[] bmText : bookmarksText) {
      final String bookmarkName = bmText[0];
      final String bookmarkValue = bmText[1];
      final boolean exist = Dispatch.call(bookmarksDispatch, "Exists", bookmarkName).getBoolean();
      if (exist) {
        final Dispatch item = Dispatch.call(bookmarksDispatch, "item", bookmarkName).toDispatch();
        final Dispatch range = Dispatch.call(item, "Range").toDispatch();
        // 设置 Text 属性的效果相当于在书签后插入文本, 即调用 InsertAfter/InsertBefore(?)
        Dispatch.put(range, "Text", new Variant(bookmarkValue));
        // Dispatch.call(range, "InsertAfter", new Variant(bookmarkValue));
        // Dispatch.call(range, "InsertBefore", new Variant(bookmarkValue));
      }
    }
  }

  /**
   * 在指定书签处的内部插入文本
   *
   * 参考 {link: https://gregmaxey.com/word_tip_pages/insert_text_at_or_in_bookmark.html}
   *
   * TODO: 如果指定书签有缺失，怎么办？
   *
   * @param bookmarksText
   *          待插入的书签文本列表
   * @return 实际插入的书签值的数量
   */
  public int insertInBookmarksByList(final List<String[]> bookmarksText) {
    int count = bookmarksText.size();
    final Dispatch bookmarksDispatch = getBookmarksDispatch();
    for (final String[] bmText : bookmarksText) {
      final String bookmarkName = bmText[0];
      final String bookmarkValue = bmText[1];
      // 书签是否存在
      final boolean exist = Dispatch.call(bookmarksDispatch, "Exists", bookmarkName).getBoolean();
      if (exist) {
        // 取书签所在的 range
        final Dispatch item = Dispatch.call(bookmarksDispatch, "item", bookmarkName).toDispatch();
        final Dispatch range = Dispatch.call(item, "Range").toDispatch();

        // 设置 Text 属性
        Dispatch.put(range, "Text", new Variant(bookmarkValue));

        // 新加入这个书签, 效果相当于删除原书签, 在原位置处加入新书签, 新书签的范围覆盖了书签值文本
        Dispatch.call(bookmarksDispatch, "Add", bookmarkName, range).toDispatch();
      } else {
        count--;
      }
    }

    return count;
  }

  /**
   * 在指定书签处的内部插入文本
   *
   * @param map
   *          参数表
   * @return 实际插入的书签值的数量
   */
  public int insertInBookmarks(final Map<String, String> map) {
    int count = map.size();
    final Dispatch bookmarksDispatch = getBookmarksDispatch();
    for (final Map.Entry<String, String> entry : map.entrySet()) {
      final String bookmarkName = entry.getKey();
      final String bookmarkValue = entry.getValue();
      // 书签是否存在
      final boolean exist = Dispatch.call(bookmarksDispatch, "Exists", bookmarkName).getBoolean();
      if (exist) {
        // 取书签所在的 range
        final Dispatch item = Dispatch.call(bookmarksDispatch, "item", bookmarkName).toDispatch();
        final Dispatch range = Dispatch.call(item, "Range").toDispatch();

        // TODO: 通过参数控制在书签处附加文本，或替换文本

        // 设置 Text 属性
        Dispatch.put(range, "Text", new Variant(bookmarkValue));

        // 新加入这个书签, 效果相当于删除原书签, 在原位置处加入新书签, 新书签的范围覆盖了书签值文本
        Dispatch.call(bookmarksDispatch, "Add", bookmarkName, range).toDispatch();
      } else {
        count--;
      }
    }

    return count;
  }

  /**
   * 开始为word文档添加水印
   *
   * @param wordPath
   *          word文档的路径
   * @param waterMarkPath
   *          添加的水印图片路径
   * @param waterMarkPath
   *          返回文档路径
   * @param msWordFormatint
   *          文档类型
   * @param left abscissa
   *          水印横坐标
   * @param top ordinate
   *          水印纵坐标
   * @return 是否成功添加
   */
  public boolean addWaterMark(final Map<String, String> map) {
    final String abscissa = map.get("abscissa");
    final String ordinate = map.get("ordinate");
    final String waterMarkCode = map.get("waterMarkCode");
    final String waterMarkName = map.get("waterMarkName");

    AppProperties appProperties = new AppProperties();
//    String picPath = appProperties.getTempDocPath() + File.separator +"1.png";
    String picPath = new File(appProperties.getPicturePath()).getAbsolutePath();
//    boolean baseToImg = Base64Util.Base64ToImage(waterMarkCode, picPath);
//    if (baseToImg) {
//      throw new RuntimeException("Error taking addWaterMark, switch base64 to Img");
//    }

    try {
//          final Dispatch docSelection = Dispatch.get(document, "Selection").toDispatch();
        final Dispatch docSelection = objWord.getProperty("Selection").toDispatch();
      // 声明word文档当前活动视窗对象
      final Dispatch activeWindow = objWord.getProperty("ActiveWindow").toDispatch();
      // 取得活动窗格对象
      Dispatch activePan = Dispatch.get(activeWindow, "ActivePane").toDispatch();
      // 取得视窗对象
      Dispatch view = Dispatch.get(activePan, "View").toDispatch();
      // 打开页眉，值为9，页脚为10
      Dispatch.put(view, "SeekView", new Variant(9));
      // 获取页眉和页脚
      Dispatch headfooter = Dispatch.get(docSelection, "HeaderFooter").toDispatch();
      // 获取水印图形对象
      Dispatch shapes = Dispatch.get(headfooter, "Shapes").toDispatch();
      // 给文档全部加上水印,设置了水印效果，内容，字体，大小，是否加粗，是否斜体，左边距，上边距。
      // 调用shapes对象的AddPicture方法将全路径为picname的图片插入当前文档
      Dispatch picture = Dispatch.call(shapes, "AddPicture", picPath).toDispatch();
      // 选择当前word文档的水印
      Dispatch.call(picture, "Select");
      Dispatch.put(picture, "Left", new Variant(abscissa));
      Dispatch.put(picture, "Top", new Variant(ordinate));
      Dispatch.put(picture, "Width", new Variant(500));
      Dispatch.put(picture, "Height", new Variant(500));

      // 关闭页眉
      Dispatch.put(view, "SeekView", new Variant(0));

      return true;
    } catch (Exception e) {
      e.printStackTrace();
      return false;
    }
  }

}