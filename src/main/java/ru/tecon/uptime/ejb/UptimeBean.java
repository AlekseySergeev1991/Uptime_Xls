package ru.tecon.uptime.ejb;

import jakarta.annotation.Resource;
import jakarta.ejb.Stateless;
import ru.tecon.uptime.Uptime;

import javax.sql.DataSource;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.logging.Logger;

@Stateless
public class UptimeBean {

    private static final Logger LOGGER = Logger.getLogger(UptimeBean.class.getName());

    @Resource(name = "jdbc/DataSourceR")
    private DataSource dsR;
    @Resource(name = "jdbc/DataSource")
    private DataSource dsRW;


    public void createReport(int reportId) {
        ExecutorService executor = Executors.newSingleThreadExecutor();
        executor.execute(() -> {
            try {
                Uptime.makeReport(reportId, dsR, dsRW);
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }
        });
    }

}
