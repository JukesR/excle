package com.excel.test;

import com.google.common.collect.Lists;
import org.junit.Ignore;
import org.junit.Test;

/**
 * Created by smt2 on 16-10-8.
 */
public class ExcleTest {


    @Test
    @Ignore
    public void writeText() {
        Mapper mapper = new Mapper();
        try {
            Bean bean = new Bean();
            bean.setName("1");
            bean.setAge(2);
            mapper.write("/tmp/",Lists.newArrayList(bean));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    @Ignore
    public void readText() {
        Mapper mapper = new Mapper();
        try {
            Bean bean = new Bean();
            bean.setName("1");
            bean.setAge(2);
            mapper.write("/tmp/",Lists.newArrayList(bean));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
