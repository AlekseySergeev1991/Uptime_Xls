package ru.tecon.uptime.model;

import java.io.Serializable;
import java.time.LocalDateTime;
import java.util.StringJoiner;

/**
 * Класс описывающий тип отчета с указанием дат и интервалов
 * @author Aleksey Sergeev
 */
public class RepType implements Serializable {
    private LocalDateTime beg;
    private LocalDateTime end;

    public RepType() {
    }

    public RepType(LocalDateTime beg, LocalDateTime end) {
        this.beg = beg;
        this.end = end;
    }

    public LocalDateTime getBeg() {
        return beg;
    }

    public void setBeg(LocalDateTime beg) {
        this.beg = beg;
    }

    public LocalDateTime getEnd() {
        return end;
    }

    public void setEnd(LocalDateTime end) {
        this.end = end;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", RepType.class.getSimpleName() + "[", "]")
                .add("beg=" + beg)
                .add("end=" + end)
                .toString();
    }
}
