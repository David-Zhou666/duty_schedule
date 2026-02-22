package com.schedule.model;

import java.util.Date;

/**
 * 节假日模型
 */
public class Holiday {
    private Date date;
    private String name;

    public Holiday() {
    }

    public Holiday(Date date, String name) {
        this.date = date;
        this.name = name;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
