package org.xli.excel.entity;

/**
 * @author xli
 * @Description
 * @Date 创建于 2019/1/23 18:29
 */
public class WinningStudent {
    private int serialNum;
    private String item;
    private String name;
    private String className;

    public WinningStudent() {
    }

    public String getClassName() {
        return className;
    }

    public void setClassName(String className) {
        this.className = className;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getItem() {
        return item;
    }

    public void setItem(String item) {
        this.item = item;
    }

    public int getSerialNum() {
        return serialNum;
    }

    public void setSerialNum(int serialNum) {
        this.serialNum = serialNum;
    }

    @Override
    public String toString() {
        return "WinningStudent{" +
                "serialNum=" + serialNum +
                ", item='" + item + '\'' +
                ", name='" + name + '\'' +
                ", className='" + className + '\'' +
                '}';
    }
}
