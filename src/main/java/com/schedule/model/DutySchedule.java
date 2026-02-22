package com.schedule.model;

import java.util.Date;
import java.util.List;

/**
 * 排班信息模型
 */
public class DutySchedule {
    private Date date;                    // 日期
    private boolean isHoliday;            // 是否节假日
    private List<Person> maleDuty;        // 男生值班人员（4人）
    private List<Person> femaleDuty;      // 女生值班人员（2人）

    public DutySchedule() {
    }

    public DutySchedule(Date date, boolean isHoliday, List<Person> maleDuty, List<Person> femaleDuty) {
        this.date = date;
        this.isHoliday = isHoliday;
        this.maleDuty = maleDuty;
        this.femaleDuty = femaleDuty;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public boolean isHoliday() {
        return isHoliday;
    }

    public void setHoliday(boolean holiday) {
        isHoliday = holiday;
    }

    public List<Person> getMaleDuty() {
        return maleDuty;
    }

    public void setMaleDuty(List<Person> maleDuty) {
        this.maleDuty = maleDuty;
    }

    public List<Person> getFemaleDuty() {
        return femaleDuty;
    }

    public void setFemaleDuty(List<Person> femaleDuty) {
        this.femaleDuty = femaleDuty;
    }
}
