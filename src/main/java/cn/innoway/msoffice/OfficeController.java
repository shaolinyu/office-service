
package cn.innoway.msoffice;

import cn.innoway.application.AppProperties;
import cn.innoway.utils.Utils;
import com.google.common.io.ByteStreams;
import com.google.common.io.Files;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Map;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.function.BiConsumer;

@RestController
public class OfficeController {
  private static final Logger logger = LoggerFactory.getLogger(OfficeController.class);

  private static final Object lock = new Object();

  // 临时文件按日期存放，日期用 datePath 保存
  private static String datePath = "19990909";

  // WORD服务队列，成员数量在启动时根据属性确定
  private static  BlockingQueue<MsWordService> serviceQueue;
  // excel服务队列，成员数量在启动时根据属性确定
  private static  BlockingQueue<MsExcelService> excelQueue;
  // ppt服务队列，成员数量在启动时根据属性确定
  private static  BlockingQueue<MsPPtService> pptQueue;

  // 临时文档的绝对目录
  private final String tempDocAbsPath;

  private final boolean showWordWindow;

  private final String waterMarkPath;

  private final DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd");

  @Autowired
  public OfficeController(final AppProperties appProperties) {
    tempDocAbsPath = new File(appProperties.getTempDocPath()).getAbsolutePath();
    showWordWindow = appProperties.getShowWordWindow();
    waterMarkPath = appProperties.getPicturePath();
    initServiceQueue(appProperties.getWordInstanceCount());
    initExcelQueue(appProperties.getWordInstanceCount());
    initPPtQueue(appProperties.getWordInstanceCount());
  }

  /**
   * 初始化服务队列
   *
   * @param size
   *          队列容量
   */
  private void initServiceQueue(final int size) {
    synchronized (lock) {
      if (serviceQueue == null) {
        serviceQueue = new ArrayBlockingQueue<>(size, true);
        // 生成服务对象
        for (int i = 0; i < size; i++) {
          if (!serviceQueue.offer(new MsWordService())) {
            throw new RuntimeException("Impossible! Unable to add MsWordService into queue.");
          }
        }

        logger.info("==> # of word instances: {}", serviceQueue.size());
      }
    }
  }

  private void initExcelQueue(final int size) {
        synchronized (lock) {
            if (excelQueue == null) {
                excelQueue = new ArrayBlockingQueue<>(size, true);
                // 生成服务对象
                for (int i = 0; i < size; i++) {
                    if (!excelQueue.offer(new MsExcelService())) {
                        throw new RuntimeException("Impossible! Unable to add MsExcelService into queue.");
                    }
                }
                logger.info("==> # of excel instances: {}", serviceQueue.size());
            }
        }
    }

  private void initPPtQueue(final int size) {
        synchronized (lock) {
            if (pptQueue == null) {
                pptQueue = new ArrayBlockingQueue<>(size, true);
                // 生成服务对象
                for (int i = 0; i < size; i++) {
                    if (!pptQueue.offer(new MsPPtService())) {
                        throw new RuntimeException("Impossible! Unable to add MsPPtService into queue.");
                    }
                }
                logger.info("==> # of ppt instances: {}", pptQueue.size());
            }
        }
    }

  private MsWordService takeService() {
    for (int tries = 3; tries > 0; tries--) {
      try {
        // 从队列中取服务，若队列空将等待
        final MsWordService msWordService = serviceQueue.take();
        return msWordService;
      } catch (final InterruptedException e) {
        // 等待时被中断，记录日志，再重试
        logger.warn("takeService interrupted: {}", e.getCause());
      }
    }
    return null;
  }

  private void putService(final MsWordService msWordService) {
    for (int tries = 3; tries > 0; tries--) {
      try {
        serviceQueue.put(msWordService);
        return;
      } catch (final InterruptedException e) {
        // 等待时被中断，记录日志，再重试
        logger.warn("putService interrupted: {}", e.getCause());
      }
    }
  }

  private MsExcelService takeExlService() {
      for (int tries = 3; tries > 0; tries--) {
        try {
            // 从队列中取服务，若队列空将等待
            final MsExcelService excelService = excelQueue.take();
            return excelService;
        } catch (final InterruptedException e) {
            // 等待时被中断，记录日志，再重试
            logger.warn("takeService interrupted: {}", e.getCause());
        }
      }
      return null;
  }

  private void putExlService(final MsExcelService exlService) {
        for (int tries = 3; tries > 0; tries--) {
            try {
                excelQueue.put(exlService);
                return;
            } catch (final InterruptedException e) {
                // 等待时被中断，记录日志，再重试
                logger.warn("putService interrupted: {}", e.getCause());
            }
        }
    }

  private MsPPtService takePPtService() {
        for (int tries = 3; tries > 0; tries--) {
            try {
                // 从队列中取服务，若队列空将等待
                MsPPtService pptService = pptQueue.take();
                return pptService;
            } catch (final InterruptedException e) {
                // 等待时被中断，记录日志，再重试
                logger.warn("takeService interrupted: {}", e.getCause());
            }
        }
        return null;
    }

  private void putPPtService(final MsPPtService pptService) {
    for (int tries = 3; tries > 0; tries--) {
        try {
            pptQueue.put(pptService);
            return;
        } catch (final InterruptedException e) {
            // 等待时被中断，记录日志，再重试
            logger.warn("putService interrupted: {}", e.getCause());
        }
    }
  }

  /**
   * 调用Word服务操作文档
   *
   * @param msWordService
   *          Word文档服务
   * @param wordProcessing
   *          文档操作内容
   * @param sourceFilePath
   *          源文件绝对路径
   * @param targetFilePath
   *          目录文件绝对路径
   * @param msWordFormat
   *          结果文档的格式
   */
  protected void callWordService(final MsWordService msWordService,
      final WordProcessing wordProcessing, final String sourceFilePath, final String targetFilePath,
      final MsWordFormat msWordFormat) {
    logger.info(Thread.currentThread().getName()+"\t start callWordService");
    // 启动Word
    msWordService.startWord(showWordWindow);
    // 打开文件
    msWordService.open(sourceFilePath);
    // 操作文件
    for (final WordProcessing.Action action : wordProcessing.getActions()) {
      final BiConsumer<MsWordService, Map<String, String>> method = MsWordService
          .getMethod(action.getMethod());
      if (method == null) {
        continue;
      }
      method.accept(msWordService, action.getArgs());
    }
    // 另存操作结果
    // TODO: 用 exportas , 可加选项参数，选择 PDF-A 格式, 参见 {@link:
    // https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb256835(v=office.12)?redirectedfrom=MSDN
    // }
    msWordService.saveAsAndClose(targetFilePath, msWordFormat);
  }

    /**
     * 调用excel服务类
     * @param excelService
     * @param wordProcessing
     * @param sourceFilePath 原路径
     * @param targetFilePath 目标路径
     */
  public void callExcelService(final  MsExcelService excelService,
                               final WordProcessing wordProcessing,
                               final String sourceFilePath,
                               final String targetFilePath){
      excelService.OpenExcel(sourceFilePath, null);
      for (final WordProcessing.Action action: wordProcessing.getActions()) {
          final  BiConsumer<MsExcelService, Map<String, String>> method =
                  MsExcelService.getMethod(action.getMethod());
          if (method == null){
              continue;
          }
          method.accept(excelService, action.getArgs());
      }
      excelService.CloseExcel(targetFilePath);
  }

    /**
     * 调用ppt服务类
     * @param wordProcessing
     * @param sourceFilePath 原路径
     * @param targetFilePath 目标路径
     */
    public void callPPtService(final MsPPtService pPtService,
                                 final WordProcessing wordProcessing,
                                 final String sourceFilePath,
                                 final String targetFilePath){
        pPtService.openDocument(sourceFilePath);
        for (final WordProcessing.Action action: wordProcessing.getActions()) {
            final  BiConsumer<MsPPtService, Map<String, String>> method =
                    MsPPtService.getMethod(action.getMethod());
            if (method == null){
                continue;
            }
            method.accept(pPtService, action.getArgs());
        }
        pPtService.closeAndSavePpt(targetFilePath);
    }
    /**
     * ppt操作水印
     */
  @PostMapping("/ppt-documents")
  public  void processPPtDocument(@RequestParam("processing") final WordProcessing wordProcessing,
                                  @RequestPart("file") final MultipartFile file,
                                  final HttpServletRequest request,
                                  final HttpServletResponse response) throws IOException {

      logger.debug("开始处理PPT*****************");
      final long begin = System.currentTimeMillis();
      //文件校验
      fileCommon(wordProcessing, file);
      String targetFileName = wordProcessing.getTargetFileName();
      // 文件存于临时文档目录下，按日期分子目录，文件名前缀以JVM时间
      final String tempPrefix = String.valueOf(System.nanoTime());
      final String sourceFilePath = tempDocAbsPath + File.separator + datePath + File.separator
              + tempPrefix + "_" + file.getOriginalFilename();
      final String targetFilePath = tempDocAbsPath + File.separator + datePath + File.separator
              + tempPrefix + "_" + targetFileName;

      final Thread c = Thread.currentThread();

      // 保存文件到临时目录，若保存出错抛出 IOException
      file.transferTo(Paths.get(sourceFilePath));

      // 从队列中取WORD服务
      final MsPPtService pPtService = takePPtService();

      logger.debug("==> {} ({}) before serviceQueue.take", c.getName(), c.getId());

      callPPtService(pPtService, wordProcessing, sourceFilePath, targetFilePath);

      putPPtService(pPtService);

      logger.debug("==> {} ({}) msWordService put", c.getName(), c.getId());
      logger.debug("==> {} converted to {} in {} ms", file.getOriginalFilename(),
              Files.getFileExtension(targetFileName), System.currentTimeMillis() - begin);

       // 回送文件
      serveResource(targetFilePath, targetFileName, request, response);
  }

    /**
     * excel 文档操作
     */
  @PostMapping("/excel-documents")
  public void processExcelDocument(@RequestParam("processing") final WordProcessing wordProcessing,
                                   @RequestPart("file") final MultipartFile file,
                                   final HttpServletRequest request,
                                   final HttpServletResponse response) throws IOException {

      logger.debug("开始处理Excel*****************");
      final long begin = System.currentTimeMillis();
      //文件校验
      fileCommon(wordProcessing, file);
      String targetFileName = wordProcessing.getTargetFileName();
      // 文件存于临时文档目录下，按日期分子目录，文件名前缀以JVM时间
      final String tempPrefix = String.valueOf(System.nanoTime());
      final String sourceFilePath = tempDocAbsPath + File.separator + datePath + File.separator
              + tempPrefix + "_" + file.getOriginalFilename();
      final String targetFilePath = tempDocAbsPath + File.separator + datePath + File.separator
              + tempPrefix + "_" + targetFileName;

      final Thread c = Thread.currentThread();

      // 保存文件到临时目录，若保存出错抛出 IOException
      file.transferTo(Paths.get(sourceFilePath));

      logger.debug("==> {} ({}) before serviceQueue.take", c.getName(), c.getId());

      // 从队列中取WORD服务
      final MsExcelService xlsService = takeExlService();
      if (xlsService == null) {
          throw new RuntimeException("Error taking service, maybe interrupted too often");
      }

      logger.debug("==> {} ({}) offerService taken", c.getName(), c.getId());

      callExcelService(xlsService, wordProcessing, sourceFilePath, targetFilePath);

      // 把服务放回到队列中
      putExlService(xlsService);

      logger.debug("==> {} ({}) msWordService put", c.getName(), c.getId());
      logger.debug("==> {} converted to {} in {} ms", file.getOriginalFilename(),
              Files.getFileExtension(targetFileName), System.currentTimeMillis() - begin);

      // 回送文件
      serveResource(targetFilePath, targetFileName, request, response);
  }

  /**
   * WORD文档操作
   *
   * @param wordProcessing
   *          WORD文档处理的动作
   * @param file
   *          待处理的源文件
   */
    @PostMapping("/word-documents")
    public void processWordDocument(@RequestParam("processing") final WordProcessing wordProcessing,
                                    @RequestPart("file") final MultipartFile file,
                                    final HttpServletRequest request,
                                    final HttpServletResponse response) throws IOException {
        logger.debug("开始处理Word*****************");

        final long begin = System.currentTimeMillis();
        final LocalDate now = LocalDate.now();
        // 请求有效性检查
        if (wordProcessing == null) {
          throw new IllegalArgumentException("No part of processing");
        }

        final String targetFileName = wordProcessing.getTargetFileName();
        if (targetFileName.isEmpty()) {
          throw new IllegalArgumentException("No targetFileName of wordProcessing");
        }

        // 根据后缀名确定目标文件格式
        final MsWordFormat msWordFormat = MsWordFormat.of(Files.getFileExtension(targetFileName));
        if (msWordFormat == null) {
          throw new IllegalArgumentException(
              "Invalid or unknown extension for file: " + targetFileName);
        }

        if (file.isEmpty()) {
          throw new IllegalArgumentException("No file body");
        }

        if (!MsWordService.isSourceFileFormatSupported(file.getOriginalFilename())) {
          throw new IllegalArgumentException("Source file format not allowed");
        }

        // 看时间是否滑到下一天
        final String today = formatter.format(now);
        if (!today.equals(datePath)) {
          synchronized (lock) {
            if (!today.equals(datePath)) {
              // 创建当天的临时目录
              Utils.checkOrCreateDir(tempDocAbsPath + File.separator + today);
              datePath = today;
            }
          }
        }
        // 文件存于临时文档目录下，按日期分子目录，文件名前缀以JVM时间
        final String tempPrefix = String.valueOf(System.nanoTime());
        final String sourceFilePath = tempDocAbsPath + File.separator + datePath + File.separator
            + tempPrefix + "_" + file.getOriginalFilename();
        final String targetFilePath = tempDocAbsPath + File.separator + datePath + File.separator
            + tempPrefix + "_" + targetFileName;

        final Thread c = Thread.currentThread();

        // 保存文件到临时目录，若保存出错抛出 IOException
        file.transferTo(Paths.get(sourceFilePath));

        logger.debug("==> {} ({}) before serviceQueue.take", c.getName(), c.getId());

        // 从队列中取WORD服务
        final MsWordService msWordService = takeService();
        if (msWordService == null) {
          throw new RuntimeException("Error taking service, maybe interrupted too often");
        }

        logger.debug("==> {} ({}) msWordService taken", c.getName(), c.getId());

        callWordService(msWordService, wordProcessing, sourceFilePath, targetFilePath, msWordFormat);

        // 把服务放回到队列中
        putService(msWordService);

        logger.debug("==> {} ({}) msWordService put", c.getName(), c.getId());
        logger.debug("==> {} converted to {} in {} ms", file.getOriginalFilename(),
            Files.getFileExtension(targetFileName), System.currentTimeMillis() - begin);

        // 回送文件
        serveResource(targetFilePath, targetFileName, request, response);
  }

  protected void serveResource(final String filePath, final String fileName,
      final HttpServletRequest request, final HttpServletResponse response) {
    final File file = new File(filePath);
    if (!file.exists()) {
      throw new RuntimeException("File '" + filePath + "' not found");
    }

    final long fileLength = file.length();
    if ((fileLength == 0) || !file.canRead()) {
      throw new RuntimeException("File '" + filePath + "' cannot be read");
    }

    try (final InputStream in = new FileInputStream(file)) {
      response.reset();
      final String mimeType = request.getServletContext().getMimeType(filePath);
      response.setHeader("Content-Type", mimeType);
      response.setHeader("content-disposition",
          "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
      response.setContentLength((int) fileLength);

      ByteStreams.copy(in, response.getOutputStream());
    } catch (final UnsupportedEncodingException e) {
      // 这个不可能，忽略了
        logger.error("UnsupportedEncodingException:{}",e);
    } catch (final IOException e) {
      // 可能是客户端关掉了连接
      logger.warn("Oops! {}", e.getCause());
    }
  }

  /**
   * 下载文档
   *
   */
  @GetMapping("/word-documents/{id}")
  public void downloadFile(@PathVariable("id") final String fileName,
      final HttpServletRequest request, final HttpServletResponse response) {

    final Thread c = Thread.currentThread();
    // logger.debug("==> {} ({}), object hascode = {}, semaphore hashcode = {}", c.getName(),
    // c.getId(), this.hashCode(), semaphore.hashCode());

    try {

      // WORD文件操作，加锁
      logger.debug("==> {} ({}) before semaphore.acquire", c.getName(), c.getId());
      // semaphore.acquireUninterruptibly();
      logger.debug("==> {} ({}) semaphore acquired", c.getName(), c.getId());

      c.sleep(15 * 1000);

    } catch (final InterruptedException e) {
      throw new RuntimeException(e.getCause());
    } finally {
      // semaphore.release();
      // logger.debug("==> {} ({}) semaphore released. availablePermits = {}", c.getName(),
      // c.getId(),
      // semaphore.availablePermits());
    }

  }

    /**
     * 校验文件
     * @param wordProcessing
     * @param file
     * @throws IOException
     */
  public void fileCommon(WordProcessing wordProcessing, MultipartFile file) throws IOException {
      final LocalDate now = LocalDate.now();
      // 请求有效性检查
      if (wordProcessing == null) {
          throw new IllegalArgumentException("No part of processing");
      }
      final String targetFileName = wordProcessing.getTargetFileName();
      if (targetFileName.isEmpty()) {
          throw new IllegalArgumentException("No targetFileName of wordProcessing");
      }
      // 根据后缀名确定目标文件格式
      final MsWordFormat msWordFormat = MsWordFormat.of(Files.getFileExtension(targetFileName));
      if (msWordFormat == null) {
          throw new IllegalArgumentException(
                  "Invalid or unknown extension for file: " + targetFileName);
      }
      if (file.isEmpty()) {
          throw new IllegalArgumentException("No file body");
      }
      if (!MsWordService.isSourceFileFormatSupported(file.getOriginalFilename())) {
          throw new IllegalArgumentException("Source file format not allowed");
      }
      // 看时间是否滑到下一天
      final String today = formatter.format(now);
      if (!today.equals(datePath)) {
          synchronized (lock) {
              if (!today.equals(datePath)) {
                  // 创建当天的临时目录
                  Utils.checkOrCreateDir(tempDocAbsPath + File.separator + today);
                  datePath = today;
              }
          }
      }
  }

}
