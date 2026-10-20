package ru.tecon.uptime;

import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.tecon.uptime.model.*;

import javax.sql.DataSource;
import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.sql.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.Date;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Collectors;

public class Uptime {

    private static final Logger LOGGER = Logger.getLogger(Uptime.class.getName());
    private static final String LOAD_REP_TYPE = "SELECT * FROM dsp_0104t.get_rep_info(?)";
    private static final String LOAD_OBJECT = "SELECT * FROM dsp_0104t.sel_object_list(?);";
    private static final String GET_PARAM_COUNT = "select dsp_0104t.get_obj_par_count(?,?)";
    private static final String LOAD_PARAMS = "SELECT * FROM dsp_0104t.get_common_obj_par_values (?, ?, ?, ?)";
    private static final String PERCENT = "call dsp_0104t.update_percent(?, ?)";
    private static final String INTERRUPTED = "SELECT * FROM dsp_0104t.get_rep_status(?)";
    private static final String SQL = "select dsp_0104t.save_report(?,?)";
    private static final String FINSQL = "call dsp_0104t.update_status(?, ?)";

    private DataSource dsR;
    public void setDsR(DataSource dsR) {
        this.dsR = dsR;
    }

    private DataSource dsRW;

    public void setDsRW(DataSource dsRW) {
        this.dsRW = dsRW;
    }



    //init метод при создании отчета
    public static void makeReport (int p_Rep_Id, DataSource dsR, DataSource dsRW) throws InterruptedException {
        long currentTime = System.nanoTime();
        LOGGER.log(Level.INFO, "start make report {0}", p_Rep_Id);

        Uptime mr = new Uptime();
        mr.setDsR(dsR);
        mr.setDsRW(dsRW);
        try {
            mr.printReport(p_Rep_Id);
        } catch (IOException | SQLException | ParseException | DecoderException e) {
            LOGGER.log(Level.WARNING, "makeReport error", e);
            e.printStackTrace();
        }

        LOGGER.log(Level.INFO, "report created {0} created time {1}", new java.lang.Object[]{p_Rep_Id, (System.nanoTime() - currentTime)});
    }

    /*
      Метод создает нужный воркбук. Параметры:
      p_Rep_Id - идентификатор отчета
      p_Rep_Type - тип репорта. Принимает на вход один символ, малая латинская буква: h - часовой, d - дневной, m - месячный
      p_Beg_Date - Начальная дата в ткстовом формате (тут пишу в ораклиной нотации) dd-mm-yyyy hh24:mi
      p_End_Date - Конечная дата в ткстовом формате (тут пишу в ораклиной нотации) dd-mm-yyyy hh24:mi. Прекрасно понимаю, что ее можно рассчитать
                   с помощью количества колонок, типа отчета и начальной даты
      p_Rows - количество строк в отчете
      p_Data_Cols - количество колонок - значений параметров. В шапке на 4 колонки больше

    */
    public void printReport (int p_Rep_Id) throws IOException, SQLException, ParseException, DecoderException, InterruptedException {
        Workbook wb2 = WorkbookFactory.create(Objects.requireNonNull(this.getClass().getClassLoader().getResourceAsStream("/Шаблон отчета.xlsx")));
        SXSSFWorkbook wb = new SXSSFWorkbook((XSSFWorkbook) wb2);
        SXSSFSheet sh = wb.createSheet("Табличный вид");
        Sheet shDia = wb2.getSheetAt(0);

        int begRow = 6;  // строка в екселе, с которой начинается собственно отчет.

        RepType repType = loadRepType(p_Rep_Id, dsRW);

        long cols = 0;
        List<LocalDateTime> dateList = new ArrayList<>();
        LocalDateTime localDateTemp = repType.getBeg();

        dateList.add(localDateTemp);
        //-- Считаем количество дней между датами
        if (repType.getEnd().getYear() == repType.getBeg().getYear() && repType.getEnd().getMonth() == repType.getBeg().getMonth()
                && repType.getEnd().getDayOfMonth() == repType.getBeg().getDayOfMonth()) {
            cols++;
        } else {
            for (;;) {
                if (cols == 0 || repType.getEnd().isAfter(localDateTemp)) {
                    if (cols != 0) {
                        localDateTemp = localDateTemp.plusDays(1);
                        dateList.add(localDateTemp);
                    }
                    cols++;
                } else {
                    break;
                }
            }
        }
            cols = cols*24;
        // Пожалуй, наполню-ка я лист отдельным методом. И сначала заполняем его, чтобы узнать Rows. Cols

        // Колонок - 4 первых и колонка с датами

        CellStyle headerStyle = setHeaderStyle(wb);
        CellStyle headerStyleNoBold = setHeaderStyleNoBold(wb);
        CellStyle nowStyle = setCellNow (wb);
        CellStyle tableHeaderStyle = setTableHeaderStyle(wb);


        // Устанавливаем ширины колонок. В конце мероприятия
        sh.setColumnWidth(0, 6 * 256);
        sh.setColumnWidth(1, 39 * 256);
        sh.setColumnWidth(2, 14 * 256);
        sh.setColumnWidth(3, 14 * 256);
        sh.setColumnWidth(4, 14 * 256);
        sh.setColumnWidth(5 + dateList.size()*24, 17 * 256);
        shDia.setColumnWidth(0, 37 * 256);
        shDia.setColumnWidth(1, 14 * 256);
        shDia.setColumnWidth(2, 21 * 256);
        shDia.setColumnWidth(3, 19 * 256);
        shDia.setColumnWidth(4, 14 * 256);


        for (int i = 5; i < cols+5; i++) {
            sh.setColumnWidth(i, 10 * 256);
            sh.setColumnHidden(i, i != 5 && i != cols + 4);
        }

        SXSSFRow row_1 = sh.createRow(0);
        row_1.setHeight((short) 435);
        SXSSFCell cell_1_1 = row_1.createCell(0);
        cell_1_1.setCellValue("ПАО \"МОЭК\": АС \"ТЕКОН - Диспетчеризация\"");
        CellRangeAddress title = new CellRangeAddress(0, 0, 0, 4);
        sh.addMergedRegion(title);
        cell_1_1.setCellStyle(headerStyle);

        SXSSFRow row_2 = sh.createRow(1);
        row_2.setHeight((short) 870);
        SXSSFCell cell_2_1 = row_2.createCell(0);
        cell_2_1.setCellValue("Статистический отчёт об общей готовности средств измерений. Табличный вид");
        CellRangeAddress formName = new CellRangeAddress(1, 1, 0, 4);
        sh.addMergedRegion(formName);
        cell_2_1.setCellStyle(headerStyle);

        // Через строчку идет период дат
        SXSSFRow row_3 = sh.createRow(2);
        row_3.setHeight((short) 435);
        SXSSFCell cell_3_1 = row_3.createCell(0);
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
        LocalDateTime begFormatted = repType.getBeg();
        String stringBeg = begFormatted.format(formatter);
        LocalDateTime endFormatted;
        endFormatted = repType.getEnd();
        String stringEnd = endFormatted.format(formatter);
        cell_3_1.setCellValue("за период: " + stringBeg + " - " + stringEnd);
        CellRangeAddress datePer = new CellRangeAddress(2, 2, 0, 4);
        sh.addMergedRegion(datePer);
        cell_3_1.setCellStyle(headerStyleNoBold);

        // Печатаем время формирования отчета
        String now = new SimpleDateFormat("dd.MM.yyyy HH:mm").format(new Date());
        SXSSFRow row_4 = sh.createRow(3);
        row_4.setHeight((short) 435);
        SXSSFCell cell_4_1 = row_4.createCell(0);
        cell_4_1.setCellStyle(nowStyle);
        cell_4_1.setCellValue("Отчет сформирован  " + now);
        CellRangeAddress nowDone = new CellRangeAddress(3, 3, 0, 4);
        sh.addMergedRegion(nowDone);

//        // Прекрасно, Заголовок сделали. Готовим шапку.
        SXSSFRow row_6 = sh.createRow(5);
        SXSSFCell cell_6_0 = row_6.createCell(0);
        cell_6_0.setCellStyle(tableHeaderStyle);
        cell_6_0.setCellValue("№");

        SXSSFCell cell_6_1 = row_6.createCell(1);
        cell_6_1.setCellStyle(tableHeaderStyle);
        cell_6_1.setCellValue("Объект/Параметр");

        SXSSFCell cell_6_2 = row_6.createCell(2);
        cell_6_2.setCellStyle(tableHeaderStyle);
        cell_6_2.setCellValue("Тех.Проц.");

        SXSSFCell cell_6_3 = row_6.createCell(3);
        cell_6_3.setCellStyle(tableHeaderStyle);
        cell_6_3.setCellValue("Ед.Изм.");

        SXSSFCell cell_6_4 = row_6.createCell(4);
        cell_6_4.setCellStyle(tableHeaderStyle);
        cell_6_4.setCellValue("Итого:");

            for (int i = 0; i < dateList.size(); i++) {
                for (int j = 0; j < 24; j++) {
                    SXSSFCell hourCell = row_6.createCell(i*24 + 5 + j);
                    hourCell.setCellStyle(tableHeaderStyle);
                    DateTimeFormatter formatterTable = DateTimeFormatter.ofPattern("dd.MM");
                    String curDateString = dateList.get(i).format(formatterTable);
                    if (j<9) {
                        hourCell.setCellValue(curDateString + " 0" + (j + 1));
                    } else if (j == 23) {
                        if (i == dateList.size() - 1) {
                            curDateString = dateList.get(i).plusDays(1).format(formatterTable);
                            hourCell.setCellValue(curDateString + " 00");
                        }else {
                            hourCell.setCellValue(curDateString + " 00");
                        }
                    }else {
                        hourCell.setCellValue(curDateString + " " + (j+1));
                    }
                }
            }

        SXSSFCell outage = row_6.createCell(5 + dateList.size()*24);
        outage.setCellStyle(tableHeaderStyle);
        outage.setCellValue("Наличие Откл");

        SXSSFCell balance = row_6.createCell(6 + dateList.size()*24);
        balance.setCellStyle(tableHeaderStyle);
        balance.setCellValue("Баланс");

        sh.createFreezePane(5, 6);

        //Повторяем упражнение для листа с диаграммами
        Row row_1Dia = shDia.createRow(0);
        row_1Dia.setHeight((short) 435);
        Cell cell_1_1Dia = row_1Dia.createCell(0);
        cell_1_1Dia.setCellValue("ПАО \"МОЭК\": АС \"ТЕКОН - Диспетчеризация\"");
        CellRangeAddress titleDia = new CellRangeAddress(0, 0, 0, 4);
        shDia.addMergedRegion(titleDia);
        cell_1_1Dia.setCellStyle(headerStyle);

        Row row_2Dia = shDia.createRow(1);
        row_2Dia.setHeight((short) 870);
        Cell cell_2_1Dia = row_2Dia.createCell(0);
        cell_2_1Dia.setCellValue("Статистический отчёт об общей готовности средств измерений. Диаграммы");
        CellRangeAddress formNameDia = new CellRangeAddress(1, 1, 0, 4);
        shDia.addMergedRegion(formNameDia);
        cell_2_1Dia.setCellStyle(headerStyle);

        // Через строчку идет период дат
        Row row_3Dia = shDia.createRow(2);
        row_3Dia.setHeight((short) 435);
        Cell cell_3_1Dia = row_3Dia.createCell(0);
        cell_3_1Dia.setCellValue("за период: " + stringBeg + " - " + stringEnd);
        CellRangeAddress datePerDia = new CellRangeAddress(2, 2, 0, 4);
        shDia.addMergedRegion(datePerDia);
        cell_3_1Dia.setCellStyle(headerStyleNoBold);

        // Печатаем время формирования отчета
        String nowDia = new SimpleDateFormat("dd.MM.yyyy HH:mm").format(new Date());
        Row row_4Dia = shDia.createRow(3);
        row_4Dia.setHeight((short) 435);
        Cell cell_4_1Dia = row_4Dia.createCell(0);
        cell_4_1Dia.setCellStyle(nowStyle);
        cell_4_1Dia.setCellValue("Отчет сформирован  " + nowDia);
        CellRangeAddress nowDoneDia = new CellRangeAddress(3, 3, 0, 4);
        shDia.addMergedRegion(nowDoneDia);

        LOGGER.log(Level.INFO, "Report head created {0}", p_Rep_Id);
        // Отлично. Заголовок и шапка сделаны. Идем по таблице, создаем и заполняем ячейки.

        LinkedList<ReportObject> objects = loadObjects(p_Rep_Id, dsR);
        DiagramValues diagramValues = createTreads(wb, p_Rep_Id, begRow, repType.getBeg(), repType.getEnd(), dateList, objects);

        fillSheetDiagram(wb2, diagramValues.getBalanceObj(), diagramValues.getNeOtklParamBalance(), diagramValues.getOtklParamBalance(),
                diagramValues.getPlanOtklParamBalance(), diagramValues.getDispOtklParamBalance(), diagramValues.getTotalAnalyzeBalance(),
                diagramValues.getTotalCausesBalance(), diagramValues.getNeBalanceObj(), diagramValues.getNeOtklParamNeBalance(),
                diagramValues.getOtklParamNeBalance(), diagramValues.getPlanOtklParamNeBalance(),
                diagramValues.getDispOtklParamNeBalance(), diagramValues.getTotalAnalyzeNeBalance(), diagramValues.getTotalCausesNeBalance());

        LOGGER.log(Level.INFO, "Report body created {0}", p_Rep_Id);


//        saveReportIntoFile(wb, "C:\\abc\\Uptime_1.xlsx");
        saveReportIntoTable (wb, p_Rep_Id, dsRW);
        wb.dispose();
    }

    private DiagramValues createTreads(SXSSFWorkbook wb, int p_Rep_Id, int begRow, LocalDateTime begDate, LocalDateTime endDate, List<LocalDateTime> dateList, LinkedList<ReportObject> objects) throws InterruptedException, DecoderException {
        DiagramValues diagramValues = new DiagramValues(0, 0, 0, 0, 0, 0,
        0, 0, 0, 0, 0, 0, 0, 0);

        int chunkSize = 100;
        AtomicInteger counter = new AtomicInteger();

        Collection<List<ReportObject>> chunkCollection = objects.stream().collect(Collectors.groupingBy(it -> counter.getAndIncrement()/chunkSize)).values();


        final BlockingQueue<ReportObject> queue = new ArrayBlockingQueue<>(chunkSize);

        int threadsCount = 10;
        objects.clear();

        BigDecimal percentage = new BigDecimal(0);
        double size = chunkCollection.size();
        double iterationPercent = 100/size;
        BigDecimal iterationPercentBD = new BigDecimal(iterationPercent);

        for (List<ReportObject> chunk : chunkCollection){
            queue.addAll(chunk);
            chunk.clear();
            ExecutorService pool = Executors.newFixedThreadPool(threadsCount);
            for (int i = 0; i < threadsCount; i++) {
                Runnable r = new Runnable() {
                    @Override
                    public void run() {
                        ReportObject reportObject;
                        while ((reportObject = queue.poll()) != null) {
                            reportObject.setParamList(loadParams(p_Rep_Id, reportObject.getObjId(), begDate, endDate, dsRW));
                            chunk.add(reportObject);
                        }
                    }
                };
                pool.execute(r);
            }
            pool.shutdown();
            if (pool.awaitTermination(10, TimeUnit.HOURS)) {
                chunk.sort(Comparator.comparing(ReportObject::getObjNum));
                for (ReportObject object : chunk){
                    if (object.getBalance().equals("Б")){
                        diagramValues.setBalanceObj(diagramValues.getBalanceObj() + 1);
                        for (Param param : object.getParamList()) {
                            if (param.getOutage().equals("N") && param.getPlan().equals("N")){
                                diagramValues.setNeOtklParamBalance(diagramValues.getNeOtklParamBalance() + 1);
                                for (Value value : param.getCurDateParam()){
                                    diagramValues.setTotalAnalyzeBalance(diagramValues.getTotalAnalyzeBalance() + 1);
                                    if (!value.getValue().equals("")){
                                        if (value.getColor() != null ){
                                            diagramValues.setTotalCausesBalance(diagramValues.getTotalCausesBalance() + 1);
                                        }
                                    }
                                }
                            } else {
                                diagramValues.setOtklParamBalance(diagramValues.getOtklParamBalance() + 1);
                                if (param.getOutage().equals("Y")){
                                    diagramValues.setDispOtklParamBalance(diagramValues.getDispOtklParamBalance() + 1);
                                }
                                if (param.getPlan().equals("Y")){
                                    diagramValues.setPlanOtklParamBalance(diagramValues.getPlanOtklParamBalance() + 1);
                                }
                                for (Value value : param.getCurDateParam()){
                                    diagramValues.setTotalAnalyzeBalance(diagramValues.getTotalAnalyzeBalance() + 1);
                                    if (!value.getValue().equals("")){
                                        if (value.getColor() != null ){
                                            diagramValues.setTotalCausesBalance(diagramValues.getTotalCausesBalance() + 1);
                                        }
                                    }
                                }
                            }
                        }
                    } else {
                        diagramValues.setNeBalanceObj(diagramValues.getNeBalanceObj() + 1);
                        for (Param param : object.getParamList()) {
                            if (param.getOutage().equals("N") && param.getPlan().equals("N")){
                                diagramValues.setNeOtklParamNeBalance(diagramValues.getNeOtklParamNeBalance() + 1);
                                for (Value value : param.getCurDateParam()){
                                    diagramValues.setTotalAnalyzeNeBalance(diagramValues.getTotalAnalyzeNeBalance() + 1);
                                    if (!value.getValue().equals("")){
                                        if (value.getColor() != null ){
                                            diagramValues.setTotalCausesNeBalance(diagramValues.getTotalCausesNeBalance() + 1);
                                        }
                                    }
                                }
                            } else {
                                diagramValues.setOtklParamNeBalance(diagramValues.getOtklParamNeBalance() + 1);
                                if (param.getOutage().equals("Y")){
                                    diagramValues.setDispOtklParamNeBalance(diagramValues.getDispOtklParamNeBalance() + 1);
                                }
                                if (param.getPlan().equals("Y")){
                                    diagramValues.setPlanOtklParamNeBalance(diagramValues.getPlanOtklParamNeBalance() + 1);
                                }
                                for (Value value : param.getCurDateParam()){
                                    diagramValues.setTotalAnalyzeNeBalance(diagramValues.getTotalAnalyzeNeBalance() + 1);
                                    if (!value.getValue().equals("")){
                                        if (value.getColor() != null ){
                                            diagramValues.setTotalCausesNeBalance(diagramValues.getTotalCausesNeBalance() + 1);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                fillSheetThreads(wb, p_Rep_Id, begRow, dateList, chunk);

                percentage = percentage.add(iterationPercentBD).setScale(3, RoundingMode.DOWN);
                percent(p_Rep_Id, percentage, dsRW);
            } else {
                LOGGER.log(Level.WARNING, "fill sheet error");
            }
        }
        return diagramValues;
    }

    private void fillSheetThreads (SXSSFWorkbook wb, int p_Rep_Id, int begRow, List<LocalDateTime> dateList, List<ReportObject> objects) throws DecoderException {
        SXSSFSheet sh = wb.getSheetAt(1);
        CellStyle cellBoldStyle = setCellBoldStyle(wb);
        CellStyle cellNoBoldStyle = setCellNoBoldStyle(wb);

        // Заполняем лист значениями, взятыми из таблицы
        if (!objects.isEmpty()) {
            HashMap<String, CellStyle> colors = new HashMap<>();

            for (ReportObject object : objects) {
                if (!interrupted(p_Rep_Id , dsR).equals("Q")) {
                    int objRow = begRow + object.getRowNum();
                    SXSSFRow row = sh.createRow(objRow);
                    SXSSFCell objNumCell = row.createCell(0);
                    objNumCell.setCellValue(object.getObjNum());
                    objNumCell.setCellStyle(cellBoldStyle);
                    SXSSFCell objNameCell = row.createCell(1);
                    objNameCell.setCellValue(object.getName());
                    objNameCell.setCellStyle(cellBoldStyle);
                    for (int i = 2; i < 5; i++) {
                        SXSSFCell cell = row.createCell(i);
                        cell.setCellStyle(cellBoldStyle);
                    }

                        for (int i = 5; i < (dateList.size()*24)+7; i++) {
                            SXSSFCell cell = row.createCell(i);
                            cell.setCellStyle(cellBoldStyle);
                        }

                    objRow++;
                    for (Param param : object.getParamList()) {
                        SXSSFRow parRow = sh.createRow(objRow);
                        SXSSFCell emptyNumCell = parRow.createCell(0);
                        emptyNumCell.setCellStyle(cellNoBoldStyle);
                        SXSSFCell parNameCell = parRow.createCell(1);
                        parNameCell.setCellValue(param.getParName());
                        parNameCell.setCellStyle(cellNoBoldStyle);
                        SXSSFCell techProcCell = parRow.createCell(2);
                        techProcCell.setCellValue(param.getTecProc());
                        techProcCell.setCellStyle(cellNoBoldStyle);
                        SXSSFCell unitCell = parRow.createCell(3);
                        unitCell.setCellValue(param.getUnits());
                        unitCell.setCellStyle(cellNoBoldStyle);
                        SXSSFCell totalCell = parRow.createCell(4);
                        totalCell.setCellValue(param.getTotal());
                        totalCell.setCellStyle(cellNoBoldStyle);
                        SXSSFCell outageCell = parRow.createCell(5 + dateList.size()*24);
                        if (param.getOutage().equals("Y") || param.getPlan().equals("Y")){
                            outageCell.setCellValue("Откл");
                        } else {
                            outageCell.setCellValue("нет Откл");
                        }
                        outageCell.setCellStyle(cellNoBoldStyle);
                        SXSSFCell balanceCell = parRow.createCell(6 + dateList.size()*24);
                        balanceCell.setCellValue(object.getBalance());
                        balanceCell.setCellStyle(cellNoBoldStyle);

                        objRow++;
                        int colNum = 0;
                        for (Value value : param.getCurDateParam()) {
                            SXSSFCell valueCell = parRow.createCell(5 + colNum);
                            valueCell.setCellValue(value.getValue());

                            if (value.getColor() != null) {
                                if (colors.containsKey(value.getColor())) {
                                    valueCell.setCellStyle(colors.get(value.getColor()));
                                } else {
                                    CellStyle cellColoredStyle = setCellNoBoldStyle(wb);
                                    String rgbS = value.getColor();
                                    byte [] rgbB = Hex.decodeHex(rgbS);
                                    XSSFColor color = new XSSFColor(rgbB, null);
                                    cellColoredStyle.setFillForegroundColor(color);
                                    cellColoredStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                                    colors.put(value.getColor(), cellColoredStyle);
                                    valueCell.setCellStyle(cellColoredStyle);
                                }
                            } else {
                                valueCell.setCellStyle(cellNoBoldStyle);
                            }
                            colNum++;
                        }
                    }
                } else {
                    break;
                }
            }
        } else {
            SXSSFRow row_9 = sh.createRow(8);
            SXSSFCell cell_9_1 = row_9.createCell(1);
            cell_9_1.setCellValue("Не выбран ни один объект");
        }
    }

    private void fillSheetDiagram (Workbook wb, int balanceObj, int neOtklParamBalance, int otklParamBalance, int planOtklParamBalance, int dispOtklParamBalance,
                                   int totalAnalyzeBalance, int totalCausesBalance, int neBalanceObj, int neOtklParamNeBalance, int otklParamNeBalance, int planOtklParamNeBalance,
                                   int dispOtklParamNeBalance, int totalAnalyzeNeBalance, int totalCausesNeBalance) throws DecoderException {
        Sheet sh = wb.getSheetAt(0);
        CellStyle invisibleStyle = setCellInvisibleStyle(wb);

        // Заполняем лист значениями, взятыми из таблицы

        Row objRow = sh.createRow(7);
        Cell objTotalNameCell = objRow.createCell(0);
        objTotalNameCell.setCellValue("Объектов в анализе:");
        Cell objTotalCell = objRow.createCell(1);
        objTotalCell.setCellValue(balanceObj + neBalanceObj);
        Cell balanceObjCell = objRow.createCell(2);
        balanceObjCell.setCellValue(balanceObj);
        Cell neBalanceObjCell = objRow.createCell(3);
        neBalanceObjCell.setCellValue(neBalanceObj);

        Row paramTotalrow = sh.createRow(8);
        Cell paramTotalNameCell = paramTotalrow.createCell(0);
        paramTotalNameCell.setCellValue("Датчиков всего:");
        Cell paramTotalCell = paramTotalrow.createCell(1);
        paramTotalCell.setCellValue(neOtklParamBalance + otklParamBalance + neOtklParamNeBalance + otklParamNeBalance);
        Cell balanceParamTotalCell = paramTotalrow.createCell(2);
        balanceParamTotalCell.setCellValue(neOtklParamBalance + otklParamBalance);
        Cell neBalanceParamTotalCell = paramTotalrow.createCell(3);
        neBalanceParamTotalCell.setCellValue(neOtklParamNeBalance + otklParamNeBalance);

        Row paramNeOtklrow = sh.createRow(9);
        Cell paramNeOtklNameCell = paramNeOtklrow.createCell(0);
        paramNeOtklNameCell.setCellValue("Датчиков без отключений:");
        Cell paramNeOtklCell = paramNeOtklrow.createCell(1);
        paramNeOtklCell.setCellValue(neOtklParamBalance + neOtklParamNeBalance);
        Cell balanceParamNeOtklCell = paramNeOtklrow.createCell(2);
        balanceParamNeOtklCell.setCellValue(neOtklParamBalance);
        Cell neBalanceParamNeOtklCell = paramNeOtklrow.createCell(3);
        neBalanceParamNeOtklCell.setCellValue(neOtklParamNeBalance);

        Row paramOtklrow = sh.createRow(10);
        Cell paramOtklNameCell = paramOtklrow.createCell(0);
        paramOtklNameCell.setCellValue("Датчиков с отключениями:");
        Cell paramOtklCell = paramOtklrow.createCell(1);
        paramOtklCell.setCellValue(otklParamBalance + otklParamNeBalance);
        Cell balanceParamOtklCell = paramOtklrow.createCell(2);
        balanceParamOtklCell.setCellValue(otklParamBalance);
        Cell neBalanceParamOtklCell = paramOtklrow.createCell(3);
        neBalanceParamOtklCell.setCellValue(otklParamNeBalance);

        Row planOtklParamrow = sh.createRow(11);
        Cell planOtklParamNameCell = planOtklParamrow.createCell(0);
        planOtklParamNameCell.setCellValue("Датчики в плановом отключении:");
        Cell planOtklParamCell = planOtklParamrow.createCell(1);
        planOtklParamCell.setCellValue(planOtklParamBalance + planOtklParamNeBalance);
        Cell planOtklParamBalanceCell = planOtklParamrow.createCell(2);
        planOtklParamBalanceCell.setCellValue(planOtklParamBalance);
        Cell planOtklParamNeBalanceCell = planOtklParamrow.createCell(3);
        planOtklParamNeBalanceCell.setCellValue(planOtklParamNeBalance);

        Row dispOtklParamrow = sh.createRow(12);
        Cell dispOtklParamNameCell = dispOtklParamrow.createCell(0);
        dispOtklParamNameCell.setCellValue("Датчики, отключенные диспетчерами:");
        Cell dispOtklParamCell = dispOtklParamrow.createCell(1);
        dispOtklParamCell.setCellValue(dispOtklParamBalance + dispOtklParamNeBalance);
        Cell dispOtklParamBalanceCell = dispOtklParamrow.createCell(2);
        dispOtklParamBalanceCell.setCellValue(dispOtklParamBalance);
        Cell dispOtklParamNeBalanceCell = dispOtklParamrow.createCell(3);
        dispOtklParamNeBalanceCell.setCellValue(dispOtklParamNeBalance);

        Row totalAnalyzerow = sh.createRow(13);
        Cell totalAnalyzeNameCell = totalAnalyzerow.createCell(0);
        totalAnalyzeNameCell.setCellValue("Общее количество измерений:");
        Cell totalAnalyzeCell = totalAnalyzerow.createCell(1);
        totalAnalyzeCell.setCellValue(totalAnalyzeBalance + totalAnalyzeNeBalance);
        Cell totalAnalyzeBalanceCell = totalAnalyzerow.createCell(2);
        totalAnalyzeBalanceCell.setCellValue(totalAnalyzeBalance);
        Cell totalAnalyzeNeBalanceCell = totalAnalyzerow.createCell(3);
        totalAnalyzeNeBalanceCell.setCellValue(totalAnalyzeNeBalance);

        Row totalCausesrow = sh.createRow(14);
        Cell totalCausesNameCell = totalCausesrow.createCell(0);
        totalCausesNameCell.setCellValue("Отклонений всего:");
        Cell totalCausesCell = totalCausesrow.createCell(1);
        totalCausesCell.setCellValue(totalCausesBalance + totalCausesNeBalance);
        Cell totalCausesBalanceCell = totalCausesrow.createCell(2);
        totalCausesBalanceCell.setCellValue(totalCausesBalance);
        Cell totalCausesNeBalanceCell = totalCausesrow.createCell(3);
        totalCausesNeBalanceCell.setCellValue(totalCausesNeBalance);

        Row uptimerow = sh.getRow(16);
        Cell uptimeBalanceCell = uptimerow.getCell(2);
        double uptimeBalanceDResult = ((double)totalAnalyzeBalance - (double)totalCausesBalance)*100/(double)totalAnalyzeBalance;
        BigDecimal uptimeBalanceResult = new BigDecimal(uptimeBalanceDResult).setScale(2, RoundingMode.HALF_UP);
        uptimeBalanceCell.setCellValue(uptimeBalanceResult.toString());

        Cell uptimeNeBalanceCell = uptimerow.getCell(3);
        double uptimeNeBalanceDResult = ((double)totalAnalyzeNeBalance - (double)totalCausesNeBalance)*100/(double)totalAnalyzeNeBalance;
        BigDecimal uptimeNeBalanceResult = new BigDecimal(uptimeNeBalanceDResult).setScale(2, RoundingMode.HALF_UP);
        uptimeNeBalanceCell.setCellValue(uptimeNeBalanceResult.toString());

        Row totalNotCausesrow = sh.createRow(17);
        Cell totalNotCausesBalanceCell = totalNotCausesrow.createCell(2);
        totalNotCausesBalanceCell.setCellValue(totalAnalyzeBalance - totalCausesBalance);
        totalNotCausesBalanceCell.setCellStyle(invisibleStyle);
        Cell totalNotCausesNeBalanceCell = totalNotCausesrow.createCell(3);
        totalNotCausesNeBalanceCell.setCellValue(totalAnalyzeNeBalance - totalCausesNeBalance);
        totalNotCausesNeBalanceCell.setCellStyle(invisibleStyle);

        Row totalDiorow = sh.getRow(21);
        Cell totalDioCell = totalDiorow.getCell(0);
        totalDioCell.setCellValue(balanceObj + neBalanceObj);

        Row topDiorow = sh.getRow(20);
        Cell topDioObjBalanceCell = topDiorow.getCell(1);
        topDioObjBalanceCell.setCellValue(balanceObj);

        Cell topDioParamBalanceCell = topDiorow.getCell(2);
        topDioParamBalanceCell.setCellValue(neOtklParamBalance + otklParamBalance);

        Cell topDioTotalBalanceCell = topDiorow.getCell(3);
        topDioTotalBalanceCell.setCellValue(totalAnalyzeBalance);

        Cell topDioUptimeBalanceCell = topDiorow.getCell(4);
        BigDecimal uptimeTop = new BigDecimal(uptimeBalanceDResult).setScale(0, RoundingMode.HALF_UP);
        topDioUptimeBalanceCell.setCellValue(uptimeTop + "%");

        Row lowDiorow = sh.getRow(22);
        Cell lowDioObjBalanceCell = lowDiorow.getCell(1);
        lowDioObjBalanceCell.setCellValue(neBalanceObj);

        Cell lowDioParamBalanceCell = lowDiorow.getCell(2);
        lowDioParamBalanceCell.setCellValue(neOtklParamNeBalance + otklParamNeBalance);

        Cell lowDioTotalBalanceCell = lowDiorow.getCell(3);
        lowDioTotalBalanceCell.setCellValue(totalAnalyzeNeBalance);

        Cell lowDioUptimeBalanceCell = lowDiorow.getCell(4);
        BigDecimal uptimeLow = new BigDecimal(uptimeNeBalanceDResult).setScale(0, RoundingMode.HALF_UP);
        lowDioUptimeBalanceCell.setCellValue(uptimeLow + "%");
    }

//    private int saveReportIntoFile (SXSSFWorkbook workbook, String file) {
//        FileOutputStream fos;
//        System.out.println("saveReportIntoFile");
//        int res = 0;
//        try {
//            fos = new FileOutputStream(file);
//            workbook.write(fos);
//
//        } catch (IOException e) {
//            e.printStackTrace();
//            System.out.println("Cant create file");
//            System.out.println(e.getMessage());
//            res = 1;
//        }
//
//        System.out.println("ok report");
//        return res;
//    }

    public RepType loadRepType(int repId, DataSource ds) {
        RepType result = new RepType();
        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_REP_TYPE)) {
            stm.setInt(1, repId);
            ResultSet res = stm.executeQuery();
            if (res.next()) {
                result.setBeg (res.getTimestamp("beg_date").toLocalDateTime());
                result.setEnd(res.getTimestamp("end_date").toLocalDateTime());
                return result;
            }

        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Rep Type", e);
        }
        return result;
    }

    public LinkedList<ReportObject> loadObjects(int repId, DataSource ds) {
        LinkedList<ReportObject> result = new LinkedList<>();
        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_OBJECT)) {
            stm.setInt(1, repId);
            ResultSet res = stm.executeQuery();
            int prevCount = 0;
            while (res.next()) {
                ReportObject item = new ReportObject(res.getInt("obj_id"), res.getString("obj_name"),
                        res.getInt("n_id"), res.getString("balance"));
                if (item.getBalance() == null){
                    item.setBalance("");
                }
                if (item.getObjNum() == 1){
                    item.setRowNum(0);
                } else {
                    item.setRowNum(result.get(result.size()-1).getRowNum() + prevCount +1);
                }

                result.add(item);
                prevCount = rowCount(repId, item.getObjId(), ds);

            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Object", e);
        }
        return result;
    }

    public int rowCount (int repId, int objId, DataSource ds) {
        int result = 0;
        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(GET_PARAM_COUNT)) {
            stm.setInt(1, repId);
            stm.setInt(2, objId);
            ResultSet res = stm.executeQuery();
            while (res.next()) {
                result = res.getInt(1);
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load rowCount", e);
        }
        return result;
    }


    public List<Param> loadParams(int repId, int objId, LocalDateTime begDate, LocalDateTime endDate, DataSource ds) {
        List<Param> result = new ArrayList<>();
        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_PARAMS)) {

            stm.setInt(1, repId);
            stm.setInt(2, objId);
            stm.setTimestamp(3, Timestamp.valueOf(begDate));
            stm.setTimestamp(4, Timestamp.valueOf(endDate));
            ResultSet res = stm.executeQuery();
            while (res.next()) {
                if (!result.isEmpty() && !result.get(result.size()-1).getStatAggr().equals(res.getString("stataggr"))) {
                    List<Value> curDateParam = new ArrayList<>();
                    Value value = new Value(res.getString("par_value"), res.getString("color"));
                    if (value.getValue() == null){
                        value.setValue("");
                    }
                    if (value.getColor()!=null) {
                        if (value.getColor().isEmpty()) {
                            value.setColor(null);
                        }
                    }

                    curDateParam.add(value);

                    Param param = new Param(null, null, null, null, null, null, null);
                    param.setParId(res.getInt("par_id"));
                    param.setParName(res.getString("par_memo"));
                    param.setTecProc(res.getString("techproc_type_char"));
                    param.setUnits(res.getString("measure_name"));
                    param.setStatAggr(res.getString("stataggr"));
                    param.setTotal(res.getString("itog"));
                    param.setCurDateParam(curDateParam);
                    param.setOutage(res.getString("otkl"));
                    param.setPlan(res.getString("plan"));

                    result.add(param);
                } else if (!result.isEmpty() && result.get(result.size()-1).getParId() == res.getInt("par_id")) {
                    Value value = new Value(res.getString("par_value"), res.getString("color"));
                    if (value.getValue() == null){
                        value.setValue("");
                    }
                    if (value.getColor() != null) {
                        if (value.getColor().isEmpty()) {
                            value.setColor(null);
                        }
                    }
                    result.get(result.size() - 1).getCurDateParam().add(value);
                } else {
                    List<Value> curDateParam = new ArrayList<>();
                    Value value = new Value(res.getString("par_value"), res.getString("color"));
                    if (value.getValue() == null){
                        value.setValue("");
                    }
                    if (value.getColor()!=null) {
                        if (value.getColor().isEmpty()) {
                            value.setColor(null);
                        }
                    }

                    curDateParam.add(value);

                    Param param = new Param(null, null, null, null, null, null, null);
                    param.setParId(res.getInt("par_id"));
                    param.setParName(res.getString("par_memo"));
                    param.setTecProc(res.getString("techproc_type_char"));
                    param.setUnits(res.getString("measure_name"));
                    param.setStatAggr(res.getString("stataggr"));
                    param.setTotal(res.getString("itog"));
                    param.setCurDateParam(curDateParam);
                    param.setOutage(res.getString("otkl"));
                    param.setPlan(res.getString("plan"));


                    result.add(param);
                }
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Params for obj " + objId, e);
        }
        return result;
    }

    public void percent(int repId, BigDecimal percent, DataSource ds) {

        try (Connection connection = ds.getConnection();
             PreparedStatement stm = connection.prepareStatement(PERCENT)) {
            stm.setInt(1, repId);
            stm.setBigDecimal(2, percent);
            stm.executeUpdate();
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "update percent error ", e);
        }
    }

    public String interrupted(int repId, DataSource ds) {
        try (Connection connection = ds.getConnection();
             PreparedStatement stm = connection.prepareStatement(INTERRUPTED)) {
            stm.setInt(1, repId);
            ResultSet res = stm.executeQuery();
            if (res.next() && (res.getString(1) != null)) {
                return res.getString(1);
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load interrupt status ", e);
        }
        return null;
    }

    public void saveReportIntoTable(SXSSFWorkbook p_wb, int repId, DataSource dsRW) {
        try (Connection connection = dsRW.getConnection();
             PreparedStatement finStmt = connection.prepareStatement(FINSQL)) {

            if  (!interrupted(repId, dsRW).equals("Q")) {
// Меняем статус и записываем блоб только если отчет не был прерван

                addBlob(p_wb, repId, dsRW);
                finStmt.setLong(1, repId);
                finStmt.setString(2, "F");

                // выполняем запись. Ура, товарищи.
                percent(repId, BigDecimal.valueOf(100),dsRW);
                finStmt.executeUpdate();
            }

        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "SQL saving blob error ", e);
        }
    }

    public void addBlob(SXSSFWorkbook wb, int repId, DataSource ds) {
        try (Connection connection = ds.getConnection();
             CallableStatement stm = connection.prepareCall(SQL);
             ByteArrayOutputStream bos = new ByteArrayOutputStream()) {

            wb.write(bos);

            stm.setLong(1, repId);
            stm.setBytes(2, bos.toByteArray());

            stm.executeQuery();
        } catch (IOException e) {
            LOGGER.log(Level.WARNING, "saving blob error ", e);
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "add blob error ", e);
        } finally {
            try {
                wb.close();
            } catch (IOException e) {
                LOGGER.log(Level.WARNING, "wb close error ", e);
            }
        }
    }

    ///////////////////////////////////////////  Определение стилей тут

    //  Стиль заголовка жирный
    private  CellStyle setHeaderStyle(SXSSFWorkbook p_wb) {

        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);

        Font headerFont = p_wb.createFont();
        headerFont.setBold(true);
        headerFont.setFontName("Times New Roman");
        headerFont.setFontHeightInPoints((short) 16);

        style.setFont(headerFont);

        return style;
    }

    //  Стиль заголовка не жирный
    private  CellStyle setHeaderStyleNoBold(SXSSFWorkbook p_wb) {

        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(false);

        Font headerFontNoBold = p_wb.createFont();
        headerFontNoBold.setBold(false);
        headerFontNoBold.setFontName("Times New Roman");
        headerFontNoBold.setFontHeightInPoints((short) 16);

        style.setFont(headerFontNoBold);

        return style;
    }

    //стиль для даты создания отчета
    private  CellStyle setCellNow(SXSSFWorkbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);


        Font nowFont = wb.createFont();
        nowFont.setBold(false);
        nowFont.setFontName("Times New Roman");
        nowFont.setFontHeightInPoints((short) 12);

        style.setFont(nowFont);

        return style;
    }

    //  Стиль шапки таблицы
    private  CellStyle setTableHeaderStyle(SXSSFWorkbook p_wb) {
        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(false);

        style.setBorderTop(BorderStyle.THICK);
        style.setBorderLeft(BorderStyle.THICK);
        style.setBorderRight(BorderStyle.THICK);
        style.setBorderBottom(BorderStyle.THICK);

        Font tableHeaderFont = p_wb.createFont();

        tableHeaderFont.setBold(true);
        tableHeaderFont.setFontName("Times New Roman");
        tableHeaderFont.setFontHeightInPoints((short) 12);

        style.setFont(tableHeaderFont);

        return style;
    }

    private  CellStyle setCellBoldStyle(SXSSFWorkbook p_wb) {
        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        style.setBorderTop(BorderStyle.THICK);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);

        Font cellBoldFont = p_wb.createFont();

        cellBoldFont.setBold(true);
        cellBoldFont.setFontName("Times New Roman");
        cellBoldFont.setFontHeightInPoints((short) 11);

        style.setFont(cellBoldFont);

        return style;
    }

    private  CellStyle setCellInvisibleStyle(Workbook p_wb) {
        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);


        Font cellBoldFont = p_wb.createFont();

        cellBoldFont.setFontName("Times New Roman");
        cellBoldFont.setFontHeightInPoints((short) 11);
        cellBoldFont.setColor(IndexedColors.WHITE.getIndex());

        style.setFont(cellBoldFont);

        return style;
    }

    private  CellStyle setCellNoBoldStyle(SXSSFWorkbook p_wb) {
        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);

        Font cellNoBoldFont = p_wb.createFont();

        cellNoBoldFont.setBold(false);
        cellNoBoldFont.setFontName("Times New Roman");
        cellNoBoldFont.setFontHeightInPoints((short) 11);

        style.setFont(cellNoBoldFont);

        return style;
    }

}
