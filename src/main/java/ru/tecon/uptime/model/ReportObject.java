package ru.tecon.uptime.model;

import java.io.Serializable;
import java.util.List;
import java.util.StringJoiner;

/**
 * Класс описывающий структуру объектов, для которых формируется отчет
 * @author Aleksey Sergeev
 */
public class ReportObject implements Serializable {

    private int objId;
    private String name;
    private List<Param> paramList;
    private int rowNum;
    private int objNum;
    private String balance;

    public ReportObject(int objId, String name, int objNum, String balance) {
        this.objId = objId;
        this.name = name;
        this.objNum = objNum;
        this.balance = balance;
    }

    public int getObjId() {
        return objId;
    }

    public void setObjId(int objId) {
        this.objId = objId;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<Param> getParamList() {
        return paramList;
    }

    public void setParamList(List<Param> paramList) {
        this.paramList = paramList;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public int getObjNum() {
        return objNum;
    }

    public void setObjNum(int objNum) {
        this.objNum = objNum;
    }

    public String getBalance() {
        return balance;
    }

    public void setBalance(String balance) {
        this.balance = balance;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", ReportObject.class.getSimpleName() + "[", "]")
                .add("objId=" + objId)
                .add("name='" + name + "'")
                .add("paramList=" + paramList)
                .add("rowNum=" + rowNum)
                .add("objNum=" + objNum)
                .add("balance='" + balance + "'")
                .toString();
    }
}
