package com.schedule.model;

/**
 * 人员信息模型
 */
public class Person {
    private String name;           // 姓名
    private String gender;         // 性别：男/女
    private boolean isInspector;   // 是否纪检委员
    private int monthlyCount;      // 本月已排班次数
    private int weeklyCount;       // 本周已排班次数

    public Person() {
    }

    public Person(String name, String gender, boolean isInspector) {
        this.name = name;
        this.gender = gender;
        this.isInspector = isInspector;
        this.monthlyCount = 0;
        this.weeklyCount = 0;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public boolean isInspector() {
        return isInspector;
    }

    public void setInspector(boolean inspector) {
        isInspector = inspector;
    }

    public int getMonthlyCount() {
        return monthlyCount;
    }

    public void setMonthlyCount(int monthlyCount) {
        this.monthlyCount = monthlyCount;
    }

    public int getWeeklyCount() {
        return weeklyCount;
    }

    public void setWeeklyCount(int weeklyCount) {
        this.weeklyCount = weeklyCount;
    }

    public void incrementMonthly() {
        this.monthlyCount++;
    }

    public void incrementWeekly() {
        this.weeklyCount++;
    }

    public void resetWeekly() {
        this.weeklyCount = 0;
    }

    @Override
    public String toString() {
        return "Person{" +
                "name='" + name + '\'' +
                ", gender='" + gender + '\'' +
                ", isInspector=" + isInspector +
                ", monthlyCount=" + monthlyCount +
                ", weeklyCount=" + weeklyCount +
                '}';
    }
}
