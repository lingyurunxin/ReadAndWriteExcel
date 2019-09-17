package com.liangpi;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * @author : 李远锋
 * @time : 2019-09-14 19:12
 * 文件说明：
 */

public class Price {

    static HashMap<String,String> priceMap = new HashMap<>();
    static {
        priceMap.put("秋","2.3,1.9");
        priceMap.put("何","2.1,1.6");
        priceMap.put("杨","2.3,1.8");
        priceMap.put("西","2.3,1.9");
        priceMap.put("棚","2,1.6");
        priceMap.put("园","2.2,1.8");
        priceMap.put("于","2.2,1.8");
        priceMap.put("新","2.3,1.8");
        priceMap.put("尹","2.2,1.7");
        priceMap.put("白","2.3,1.8");
        priceMap.put("麻","2.2,1.7");
        priceMap.put("小","2.3,1.8");
        priceMap.put("车","2.1,1.6");
        priceMap.put("七","2.3,1.9");
        priceMap.put("包二","2.3,1.9");
        priceMap.put("102","2,1.7");
        priceMap.put("邵","2,1.6");
        priceMap.put("祥","2,1.7");
        priceMap.put("正","2,1.7");
        priceMap.put("丁","1.8,1.6");
        priceMap.put("百大","1.9,1.6");
        priceMap.put("州","2.2,1.7");
        priceMap.put("好","2.2,1.7");
        priceMap.put("苑","2.3,1.8");
        priceMap.put("下","2.3,1.8");
        priceMap.put("乐","2.2,1.9");
        priceMap.put("散","2.2,1.8");
        priceMap.put("医","2.2,1.8");
        priceMap.put("汽","2.2,1.7");
        priceMap.put("物","2.3,1.8");
        priceMap.put("10区","2.3,1.9");
        priceMap.put("万","2.2,1.8");
        priceMap.put("达","2.1,1.7");
        priceMap.put("际","2.2,1.8");
        priceMap.put("九","2.2,1.8");
        priceMap.put("风","2.2,1.7");
        priceMap.put("阳","2.1,1.7");
        priceMap.put("许","1.7,1.5");
        priceMap.put("香","1.9,1.5");
        priceMap.put("福","0,0");
        priceMap.put("毛","0,0");
        priceMap.put("品","2,1.6");

    }

    public static void main(String[] args) {

        Set<Map.Entry<String, String>> entries = priceMap.entrySet();
        for (Map.Entry<String, String> entry : entries) {
            String value = entry.getValue();
            String[] split = value.split(",");
            System.out.print(entry.getKey());
            System.out.print(Double.parseDouble(split[0]));
            System.out.print(Double.parseDouble(split[1]));
            System.out.println();
        }

    }

}
