package cn.innoway.msoffice;

import com.google.common.collect.ImmutableMap;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.apache.commons.io.FileUtils;
import org.apache.tomcat.util.http.fileupload.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;

public class MsPPtService{
    private static final Logger logger = LoggerFactory.getLogger(MsPPtService.class);

    private static  final BiConsumer<MsPPtService, Map<String, String>> METHOD_PICTURE = MsPPtService::addWaterMark;

    private static  final Map<String, BiConsumer<MsPPtService, Map<String, String>>> METHODS =
            ImmutableMap.of("addWaterMark", METHOD_PICTURE);

    private ActiveXComponent ppt;
    private Dispatch presentation;

    public static BiConsumer<MsPPtService, Map<String, String>> getMethod(final  String method){
        return METHODS.get(method);
    }

    private static final List<String> methodList = Arrays.asList("addWaterMark");

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
     * 关闭PPT并释放线程
     * @throws Exception
     */
    public void closeAndSavePpt(String savePath){
        Dispatch.call(presentation, "SaveAs", savePath);//ppSaveAsPDF为特定值3
        if (null != presentation) {
            Dispatch.call(presentation, "Close");
    }
        ppt.invoke("Quit", new Variant[]{});
        ppt = null;
        ComThread.Release();
    }

    /**
     * ppt另存为
     *
     * @param presentation
     * @param saveTo
     * @param ppSaveAsFileType
     * @date 2009-7-4
     * @author YHY
     */
    public void saveAs(Dispatch presentation, String saveTo,
                       int ppSaveAsFileType)throws Exception {
        Dispatch.call(presentation, "SaveAs", saveTo, new Variant(
                ppSaveAsFileType));
    }

    /**
     * 打开PPT并且添加水印保存
     * @param filePath 源PPT文件路径
     *                  String filePath, String savePath, String content
     */
    public void  openDocument(String filePath){
        if(ppt != null){
            return;
        }
        //初始化com的线程
       ComThread.InitSTA();
       ppt = new ActiveXComponent("PowerPoint.Application");
       //ppt.setProperty("Visible", new Variant(true)); //设置可见性
       Dispatch preDispatch = ppt.getProperty("Presentations").toDispatch();
       presentation =Dispatch.call(preDispatch, "Open", filePath).toDispatch();
    }


   public boolean addWaterMark(final Map<String, String> map){
       String waterContent = map.get("waterContent");
       String size = map.get("size");
       String left = map.get("left");
       String right = map.get("right");
       //所有幻灯片
       Dispatch slides= Dispatch.get(presentation, "Slides").toDispatch();
       //获取幻灯片数量
       Variant slidesCount = Dispatch.get(slides, "Count");
       System.out.println("slidesCount:"+slidesCount);
       try {
           //遍历幻灯片
           for(int i=0;i<slidesCount.getInt();i++) {
               Dispatch slide= Dispatch.call(slides, "Item", new Variant(i+1)).toDispatch();
               //获取幻灯片内所有元素
               Dispatch shapes =   Dispatch.get(slide, "Shapes").toDispatch();
               Dispatch textEffect= Dispatch.call(shapes, "AddTextEffect",
                       new Variant(0),waterContent,"微软雅黑",
                       new Variant(size),new Variant(0),new Variant(1),
                       new Variant(left),new Variant(right)).toDispatch();
           }
       }catch (Exception e){
           e.printStackTrace();
           return false;
       }
       ComThread.quitMainSTA();
       return true;
   }

}
