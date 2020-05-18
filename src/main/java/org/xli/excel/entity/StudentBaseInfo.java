package org.xli.excel.entity;

/**
 * @author xli
 * @Description
 * @Date 创建于 2019/1/23 21:41
 */
public class StudentBaseInfo {
    private int serialNum;
    private String name;
    private String id;
    private String className;

    public String getClassName() {
        return className;
    }

    public void setClassName(String className) {
        this.className = className;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getSerialNum() {
        return serialNum;
    }

    public void setSerialNum(int serialNum) {
        this.serialNum = serialNum;
    }

    @Override
    public String toString() {
        return "StudentBaseInfo{" +
                "serialNum=" + serialNum +
                ", name='" + name + '\'' +
                ", id='" + id + '\'' +
                ", className='" + className + '\'' +
                '}';
    }
}
