package com.excel.test;

/**
 * Created by smt2 on 16-10-8.
 */
@Excel(name = "客户管理")
public class Bean {
    @Column(name = "姓名", width = 30,converter = TextConvert.class)
    String name;
    @Column(name = "年龄", width = 30)
    Integer age;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }
}
