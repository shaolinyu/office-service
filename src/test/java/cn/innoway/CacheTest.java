package cn.innoway;


import com.google.common.cache.Cache;
import com.google.common.cache.CacheBuilder;

import java.util.UUID;
import java.util.concurrent.TimeUnit;

public class CacheTest {

    public static void main(String[] args){

        String str = "hello world";
        while (true){
            str.intern();
        }

//        Cache<Object, Object> cache = CacheBuilder.newBuilder().maximumSize(11000).
//                expireAfterWrite(60, TimeUnit.SECONDS).build();
//        for (int i = 0; i < 10000; i++) {
//            cache.put(UUID.randomUUID().toString().substring(0, 8),
//                    System.currentTimeMillis());
//        }
//        System.out.println(cache.asMap());
//        try {
//            TimeUnit.SECONDS.sleep(110);
//        } catch (InterruptedException e) {
//            e.printStackTrace();
//        }
//        System.out.println(cache.asMap());

    }

}
