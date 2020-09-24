
package cn.innoway.msoffice;

/**
 * MS Word 文件格式枚举类
 * 
 */
public enum MsWordFormat {
  // for WdSaveFormat enumeration see {link:
  // https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat}
  // Microsoft Office Word 97 - 2003 binary file format.
  DOC(0, "doc"), DOCX(16, "docx"), PDF(17, "pdf"),
  XLS(1, "xls"),XLSX(18, "xlsx"),PPT(2, "ppt"), PPTX(19, "pptx");

  private final int value;

  private final String fileExtension;

  private MsWordFormat(final int value, final String fileExtension) {
    this.value = value;
    this.fileExtension = fileExtension;
  }

  /**
   * 按文件后缀取对应格式
   * 
   * @param fileExtension
   *          文件后缀
   * @return 文件格式
   */
  public static MsWordFormat of(final String fileExtension) {
    for (final MsWordFormat enumeration : MsWordFormat.values()) {
      if (enumeration.fileExtension.equals(fileExtension)) {
        return enumeration;
      }
    }

    return null;
  }

  public int getValue() {
    return value;
  }

  public String getFileExtension() {
    return fileExtension;
  }
}
