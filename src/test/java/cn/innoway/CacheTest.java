package cn.innoway;


import java.util.ArrayList;
import java.util.List;
import java.util.Map;


public class CacheTest {

    static class  OOMObject{
        public byte[] placeholder = new byte[64 * 1024];
    }


    public static void main(String[] args) throws InterruptedException {
        fillHeap(1000);
        System.gc();

        Thread.sleep(20);
    }

    public static void fillHeap(int num) throws InterruptedException {
        List<OOMObject> list = new ArrayList<>();
        for (int i = 0; i < num; i++) {
            //延迟
            Thread.sleep(50);
            list.add(new OOMObject());
        }
        Thread.sleep(20);
    }

}
