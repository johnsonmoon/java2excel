package com.xuyihao.java2excel.api.model;

import java.util.Date;

/**
 * Created by Xuyh at 2016/07/21 10:36.
 */
public class Student {
    private int id;
    private String name;
    private int age;
    private Date birth;

    public Student()
    {
    }

    public Student(int id, String name, int age, Date birth)
    {
        this.id = id;
        this.name = name;
        this.age = age;
        this.birth = birth;
    }

    public int getId()
    {
        return id;
    }

    public void setId(int id)
    {
        this.id = id;
    }

    public String getName()
    {
        return name;
    }

    public void setName(String name)
    {
        this.name = name;
    }

    public int getAge()
    {
        return age;
    }

    public void setAge(int age)
    {
        this.age = age;
    }

    public Date getBirth()
    {
        return birth;
    }

    public void setBirth(Date birth)
    {
        this.birth = birth;
    }
}
