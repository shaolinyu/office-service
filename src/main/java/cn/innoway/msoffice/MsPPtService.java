package cn.innoway.msoffice;

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

public class MsPPtService{
    private static final Logger logger = LoggerFactory.getLogger(MsPPtService.class);
   /* //临时目录
    private static final String TEMP_FILE_PATH;
    static {
        String catalinaHome=System.getProperty("catalina.home");
        if(StringUtils.isEmpty(catalinaHome)) {
            TEMP_FILE_PATH=MsPPtService.class.getResource("/").getPath()+File.separator+"temp";
            
        }else {//web环境
            TEMP_FILE_PATH=catalinaHome+File.separator+"temp";
        }
        File fileTemp=new File(TEMP_FILE_PATH);
        if(fileTemp.exists()) {
            
        }else {
            fileTemp.mkdirs();
        }
    }*/

    private ActiveXComponent ppt;
    private ActiveXComponent presentation;

    /**
     * 关闭PPT并释放线程
     * @throws Exception
     */
    public void closePpt()throws Exception{
        if (null != presentation) {
            Dispatch.call(presentation, "Close");
        }
        ppt.invoke("Quit", new Variant[]{});
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


   /* *//**
         * ppt添加水印
         * @param is 输入流
         * @param os 输出流
         * @param waterContent 水印内容
         *
         * *//*
        public static boolean addWater(InputStream is,OutputStream  os,String waterContent,String name) {
            try {
                //生成文件
                String filePath =TEMP_FILE_PATH+File.separator+name;
                File fileFrom=new File(filePath);
                OutputStream osFileFrom= FileUtils.openOutputStream(fileFrom);
                IOUtils.copy(is, osFileFrom);
                is.close();
                
                //初始化com的线程
                ComThread.InitSTA();
                //ppt程序
                ActiveXComponent ppt=null;                    
                ppt = new ActiveXComponent("PowerPoint.Application");
                Dispatch pptDocument = ppt.getProperty("Presentations").toDispatch();
                //打开文档
                Dispatch curDocument =Dispatch.call(pptDocument, "Open", filePath).toDispatch();
                
                //所有幻灯片
                Dispatch slides=    Dispatch.get(curDocument, "Slides").toDispatch();
                //获取幻灯片数量
                Variant slidesCount = Dispatch.get(slides, "Count");
                
                //遍历幻灯片
                for(int i=0;i<slidesCount.getInt();i++) {
                    Dispatch slide= Dispatch.call(slides, "Item", new Variant(i+1)).toDispatch();
                    //获取幻灯片内所有元素
                    Dispatch shapes =   Dispatch.get(slide, "Shapes").toDispatch();
                    //添加水印
                    Dispatch.call(shapes, "AddTextEffect",new Variant(0),waterContent,"宋体",new Variant(10),new Variant(0),new Variant(1),new Variant(0),new Variant(0)).toDispatch();
                }
                //保存
                Dispatch.call(curDocument, "Save");
                //关闭文件
                Dispatch.call(curDocument, "Close");
                //关闭程序
                ppt.invoke("Quit", new Variant[] {});
                //释放COM
                ComThread.quitMainSTA();
                InputStream isFileFrom=FileUtils.openInputStream(fileFrom);
                IOUtils.copy(isFileFrom, os);
                os.close();
                isFileFrom.close();
                return true;
            } catch (Exception e) {
                // TODO: handle exception
                e.printStackTrace();
            }
            return false;
            
        }*/

    /**
     * 打开PPT并且添加水印保存
     * @param filePath 源PPT文件路径
     * @param savePath  另存为PPT路径
     */
    public void openDocument(String filePath, String savePath) {
            //初始化com的线程
           ComThread.InitSTA();
//            ActiveXComponent ppt=null;
//           ActiveXComponent presentation=null;  
                 
           ppt = new ActiveXComponent("PowerPoint.Application");  
           //ppt.setProperty("Visible", new Variant(true)); //设置可见性
           Dispatch pptDocument = ppt.getProperty("Presentations").toDispatch();
           Dispatch curDocument =Dispatch.call(pptDocument, "Open", filePath).toDispatch(); 
           //所有幻灯片
           Dispatch slides= Dispatch.get(curDocument, "Slides").toDispatch();
           //获取幻灯片数量
           Variant slidesCount = Dispatch.get(slides, "Count");
           System.out.println("slidesCount:"+slidesCount);
           
           //遍历幻灯片
           for(int i=0;i<slidesCount.getInt();i++) {
               Dispatch slide= Dispatch.call(slides, "Item", new Variant(i+1)).toDispatch();
               //获取幻灯片内所有元素
               Dispatch shapes =   Dispatch.get(slide, "Shapes").toDispatch();
               Dispatch textEffect= Dispatch.call(shapes, "AddTextEffect",
                       new Variant(0),"安徽移动摩卡软件","微软雅黑",
                       new Variant(30),new Variant(0),new Variant(1),
                       new Variant(200),new Variant(400)).toDispatch();
           }
           Dispatch.call(curDocument, "SaveAs", savePath);//ppSaveAsPDF为特定值3
           Dispatch.call(curDocument, "Close");
           ComThread.quitMainSTA();
       }

}
