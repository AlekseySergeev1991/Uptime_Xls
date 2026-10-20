package ru.tecon.uptime;

import jakarta.ejb.EJB;
import jakarta.servlet.ServletException;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import ru.tecon.uptime.ejb.UptimeBean;

import java.io.IOException;

@WebServlet("/loadUptime")
public class Servlet extends HttpServlet {
    @EJB
    private UptimeBean uptimeBean;

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        int reportId = Integer.parseInt(req.getParameter("reportId"));

        uptimeBean.createReport(reportId);


        resp.setStatus(HttpServletResponse.SC_OK);
    }
}
