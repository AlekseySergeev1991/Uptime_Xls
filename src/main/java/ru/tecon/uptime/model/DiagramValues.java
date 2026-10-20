package ru.tecon.uptime.model;

import java.io.Serializable;
import java.util.StringJoiner;

public class DiagramValues implements Serializable {

    int balanceObj;
    int neOtklParamBalance;
    int otklParamBalance;
    int planOtklParamBalance;
    int dispOtklParamBalance;
    int totalAnalyzeBalance;
    int totalCausesBalance;
    int neBalanceObj;
    int neOtklParamNeBalance;
    int otklParamNeBalance;
    int planOtklParamNeBalance;
    int dispOtklParamNeBalance;
    int totalAnalyzeNeBalance;
    int totalCausesNeBalance;

    public DiagramValues(int balanceObj, int neOtklParamBalance, int otklParamBalance, int planOtklParamBalance, int dispOtklParamBalance, int totalAnalyzeBalance, int totalCausesBalance, int neBalanceObj, int neOtklParamNeBalance, int otklParamNeBalance, int planOtklParamNeBalance, int dispOtklParamNeBalance, int totalAnalyzeNeBalance, int totalCausesNeBalance) {
        this.balanceObj = balanceObj;
        this.neOtklParamBalance = neOtklParamBalance;
        this.otklParamBalance = otklParamBalance;
        this.planOtklParamBalance = planOtklParamBalance;
        this.dispOtklParamBalance = dispOtklParamBalance;
        this.totalAnalyzeBalance = totalAnalyzeBalance;
        this.totalCausesBalance = totalCausesBalance;
        this.neBalanceObj = neBalanceObj;
        this.neOtklParamNeBalance = neOtklParamNeBalance;
        this.otklParamNeBalance = otklParamNeBalance;
        this.planOtklParamNeBalance = planOtklParamNeBalance;
        this.dispOtklParamNeBalance = dispOtklParamNeBalance;
        this.totalAnalyzeNeBalance = totalAnalyzeNeBalance;
        this.totalCausesNeBalance = totalCausesNeBalance;
    }

    public int getBalanceObj() {
        return balanceObj;
    }

    public void setBalanceObj(int balanceObj) {
        this.balanceObj = balanceObj;
    }

    public int getNeOtklParamBalance() {
        return neOtklParamBalance;
    }

    public void setNeOtklParamBalance(int neOtklParamBalance) {
        this.neOtklParamBalance = neOtklParamBalance;
    }

    public int getOtklParamBalance() {
        return otklParamBalance;
    }

    public void setOtklParamBalance(int otklParamBalance) {
        this.otklParamBalance = otklParamBalance;
    }

    public int getPlanOtklParamBalance() {
        return planOtklParamBalance;
    }

    public void setPlanOtklParamBalance(int planOtklParamBalance) {
        this.planOtklParamBalance = planOtklParamBalance;
    }

    public int getDispOtklParamBalance() {
        return dispOtklParamBalance;
    }

    public void setDispOtklParamBalance(int dispOtklParamBalance) {
        this.dispOtklParamBalance = dispOtklParamBalance;
    }

    public int getTotalAnalyzeBalance() {
        return totalAnalyzeBalance;
    }

    public void setTotalAnalyzeBalance(int totalAnalyzeBalance) {
        this.totalAnalyzeBalance = totalAnalyzeBalance;
    }

    public int getTotalCausesBalance() {
        return totalCausesBalance;
    }

    public void setTotalCausesBalance(int totalCausesBalance) {
        this.totalCausesBalance = totalCausesBalance;
    }

    public int getNeBalanceObj() {
        return neBalanceObj;
    }

    public void setNeBalanceObj(int neBalanceObj) {
        this.neBalanceObj = neBalanceObj;
    }

    public int getNeOtklParamNeBalance() {
        return neOtklParamNeBalance;
    }

    public void setNeOtklParamNeBalance(int neOtklParamNeBalance) {
        this.neOtklParamNeBalance = neOtklParamNeBalance;
    }

    public int getOtklParamNeBalance() {
        return otklParamNeBalance;
    }

    public void setOtklParamNeBalance(int otklParamNeBalance) {
        this.otklParamNeBalance = otklParamNeBalance;
    }

    public int getPlanOtklParamNeBalance() {
        return planOtklParamNeBalance;
    }

    public void setPlanOtklParamNeBalance(int planOtklParamNeBalance) {
        this.planOtklParamNeBalance = planOtklParamNeBalance;
    }

    public int getDispOtklParamNeBalance() {
        return dispOtklParamNeBalance;
    }

    public void setDispOtklParamNeBalance(int dispOtklParamNeBalance) {
        this.dispOtklParamNeBalance = dispOtklParamNeBalance;
    }

    public int getTotalAnalyzeNeBalance() {
        return totalAnalyzeNeBalance;
    }

    public void setTotalAnalyzeNeBalance(int totalAnalyzeNeBalance) {
        this.totalAnalyzeNeBalance = totalAnalyzeNeBalance;
    }

    public int getTotalCausesNeBalance() {
        return totalCausesNeBalance;
    }

    public void setTotalCausesNeBalance(int totalCausesNeBalance) {
        this.totalCausesNeBalance = totalCausesNeBalance;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", DiagramValues.class.getSimpleName() + "[", "]")
                .add("balanceObj=" + balanceObj)
                .add("neOtklParamBalance=" + neOtklParamBalance)
                .add("otklParamBalance=" + otklParamBalance)
                .add("planOtklParamBalance=" + planOtklParamBalance)
                .add("dispOtklParamBalance=" + dispOtklParamBalance)
                .add("totalAnalyzeBalance=" + totalAnalyzeBalance)
                .add("totalCausesBalance=" + totalCausesBalance)
                .add("neBalanceObj=" + neBalanceObj)
                .add("neOtklParamNeBalance=" + neOtklParamNeBalance)
                .add("otklParamNeBalance=" + otklParamNeBalance)
                .add("planOtklParamNeBalance=" + planOtklParamNeBalance)
                .add("dispOtklParamNeBalance=" + dispOtklParamNeBalance)
                .add("totalAnalyzeNeBalance=" + totalAnalyzeNeBalance)
                .add("totalCausesNeBalance=" + totalCausesNeBalance)
                .toString();
    }
}
