package cn.innoway;


import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;

public class CacheTest {

    private static BlockingQueue<String> blockingQueue = new ArrayBlockingQueue<>(4);

    public static void main(String[] args){

        blockingQueue.offer("k1");
        blockingQueue.offer("k2");
        blockingQueue.offer("k3");
        blockingQueue.offer("k4");
        while (!blockingQueue.isEmpty()) {
            System.out.println(blockingQueue.poll());
        }

    }

}
