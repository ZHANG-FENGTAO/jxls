package com.example.demo.poi;



import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.util.*;

@Controller
@RequestMapping("/api/orderPrint")
public class ExportExcel {

    static Logger logger = LoggerFactory.getLogger(ExportExcel.class);
//    @Autowired
//    IOrderService orderService;
//    @Autowired
//    UPage pageUtil;
//    @Autowired
//    ISiteService siteService;


    @RequestMapping(value = "/print", produces = {"application/json;charset=UTF-8"})
    public String print(ModelMap model, HttpServletRequest request, HttpServletResponse response) throws Exception {
//        //1.创建工作簿
//        Workbook wb = new SXSSFWorkbook(100);
//        //2.创建工作表Sheet
//        Sheet sheet = wb.createSheet();
//
//        Row nRow = null;
//        Cell nCell = null;
//        int rowNo = 0;//第一行
//        int cellNo = 1;//第二列
//
//
//        //设置列宽
//
//        //========================制作大标题==========================
//        nRow = sheet.createRow(rowNo++);  //创建第一行
//        nRow.setHeightInPoints(36);//设置第一行的行高
//        nCell = nRow.createCell(cellNo);  //创建一行的第二个单元格
//        nCell.setCellStyle(UPoi.bigTitle(wb));//设置大标题样式
//        nCell.setCellValue("订单统计报表");
//        //横向合并单元格   sheet.addMergedRegion(new CellRangeAddress(开始行，结束行，开始列，结束列));
//        sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 22));
//
//
//        //日　　期	站点名称	订单编号	订单状态	订单来源
//        // 收货人名称	收货人电话	收货人地址	商家名称	商品金额  活动优惠价格
//        // 配送费用	餐盒费用	优惠券抵费用	实际支付金额	支付方式
//        // 配送员
//
//        //=========================制作小标题================================
//        nRow = sheet.createRow(rowNo++);//第二行
//        nRow.setHeightInPoints(26.25f);//设置第二行的行高
//        String titles[] = {"日期", "站点名称", "订单编号", "订单状态", "配送起始时间", "配送送达时间", "订单来源", "收货人名称", "收货人电话", "商家名称","业务员","业务员联系电话", "商品金额",
//                "活动优惠金额", "配送费用", "餐盒费用", "优惠券抵费用", "站点活动补贴", "小费", "实际支付金额", "支付方式", " 配送员"};
//        int i = 1;
//        for (String title : titles) {
//            nCell = nRow.createCell(cellNo++);//创建出相应的小标题的单元格对象
//            nCell.setCellValue(title);//设置小标题文本
//            nCell.setCellStyle(UPoi.title(wb));//设置小标题的样式
//            sheet.setColumnWidth(i++, 13 * 256);
//        }

        //===========================制作数据=====================================
        Map<String, Object> map = new HashMap<>();
        String beginMoment = null;
        // String endDate = request.getParameter("endDate");

       /* if (UValidator.isNullOrEmpty(startDate)&& UValidator.isNullOrEmpty(endDate)){
            Calendar calendar = Calendar.getInstance();
            Date now = new Date();
            calendar.setTime(now);
            calendar.add(calendar.DATE,1);
            now = calendar.getTime();
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
            endDate = sdf.format(now);

            Date dBefore = new Date();
            calendar.setTime(now);
            calendar.add(Calendar.DAY_OF_MONTH, -7);
            dBefore = calendar.getTime();
            startDate = sdf.format(dBefore);
        }*/

        String startDate = request.getParameter("startDate");
        String endDate = request.getParameter("endDate");
        String startTime = request.getParameter("startTime");
        map.put("startDate", startDate);
        map.put("endDate", endDate);

//        if (UValidator.isNullOrEmpty(startDate) && UValidator.isNullOrEmpty(endDate) && UValidator.isNullOrEmpty(startTime)) {
//            Calendar calendar = Calendar.getInstance();
//            Date now = new Date();
//            calendar.setTime(now);
//            now = calendar.getTime();
//            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
//            beginMoment = sdf.format(now);
//            map.put("beginMoment", beginMoment);
//        }
//
//        String siteId = request.getParameter("site");
//        List<SiteBean> siteList = UCast.cast(request.getSession().getAttribute("siteList"));
//        List<SiteBean> sites = siteService.getSitesForQuery(siteId, siteList);
//        map.put("sites", sites);
//        // map.put("endDate", endDate);
//
//        // List<OrderVO> list = orderService.queryOrderBeanList(pageUtil.toPageBounds(1, 1000000000), map);
//        if (!UValidator.isNullOrEmpty(startTime)) {
//            map.put("startTime", startTime + " 00:00:00");
//            map.put("endTime", startTime + " 23:59:59");
//        }
//        List<OrderVO> orderList = orderService.queryOrderAndOrderCycle(pageUtil.toPageBounds(1, 1000000000), map);
//        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
//        CellStyle cellStyle = UPoi.text(wb);
//        for (OrderVO orderVO : orderList) {
//
//            nRow = sheet.createRow(rowNo++);
//            nRow.setHeightInPoints(24);//设置行高
//            cellNo = 1;
//            //日期
//            nCell = nRow.createCell(cellNo++);
//            String str = format.format(orderVO.getAddDate());
//            nCell.setCellValue(str);
//            nCell.setCellStyle(cellStyle);
//
//            //站点名称
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(orderVO.getSiteName());
//            nCell.setCellStyle(cellStyle);
//
//            // 订单编号","类型","站点名称","实付","支付方式","支付渠道","派送员","来源","订单状态
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(orderVO.getNum());
//            nCell.setCellStyle(cellStyle);
//
//            //订单状态
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(OrderStatusEnum.DescOf(orderVO.getStatus()));
//            nCell.setCellStyle(cellStyle);
//
//            //配送起始时间
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(orderVO.getAddOrderDate());
//            nCell.setCellStyle(cellStyle);
//
//
//            //配送送达时间
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(orderVO.getSendOrderDate());
//            nCell.setCellStyle(cellStyle);
//
//            //订单来源
//            nCell = nRow.createCell(cellNo++);
//            String source = "";
//            Integer sourceT = orderVO.getSource();
//            if (!UValidator.isNullOrEmpty(sourceT)) {
//                switch (sourceT) {
//                    case 1:
//                        source = "APP";
//                        break;
//                    case 2:
//                        source = "电话";
//                        break;
//                    case 3:
//                        source = "公众号";
//                        break;
//                    case 4:
//                        source = "小程序";
//                        break;
//                    default:
//                        break;
//                }
//            }
//            nCell.setCellValue(source);
//            nCell.setCellStyle(cellStyle);
//
//            //收货人名称
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(orderVO.getToName());
//            nCell.setCellStyle(cellStyle);
//
//            //收货人电话
//            nCell = nRow.createCell(cellNo++);
//            String phone = orderVO.getToPhone();
//            if (!UValidator.isNullOrEmpty(phone)) {
//                if (phone.length() == 11) {
//                    nCell.setCellValue(phone.substring(0, 3) + "****" + phone.substring(7, phone.length()));
//                } else if(phone.length()>6){
//                    nCell.setCellValue(phone.substring(0, 3) + "****");
//                }else {
//                    nCell.setCellValue(phone);
//                }
//
//            } else {
//                nCell.setCellValue("");
//            }
//            nCell.setCellStyle(cellStyle);
//
//            //商家名称
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(orderVO.getGoodOwnName());
//            nCell.setCellStyle(cellStyle);
//
//
//            //商家业务员
//            nCell = nRow.createCell(cellNo++);
//            String counterman = orderVO.getCounterman();
//            if (UValidator.isNullOrEmpty(counterman)){
//                counterman="";
//            }
//            nCell.setCellValue(counterman);
//            nCell.setCellStyle(cellStyle);
//
//
//            //业务员联系电话
//            nCell = nRow.createCell(cellNo++);
//            String counterPhone = orderVO.getCounterPhone();
//            if (UValidator.isNullOrEmpty(counterPhone)){
//                counterPhone="";
//            }
//            nCell.setCellValue(counterPhone);
//            nCell.setCellStyle(cellStyle);
//
//
//            //商品金额
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(UString.parseString(orderVO.getPriceGood(), "0"));
//            nCell.setCellStyle(cellStyle);
//
//            //活动优惠金额
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(UString.parseString(orderVO.getPriceFavor(), "0"));
//            nCell.setCellStyle(cellStyle);
//
//            //配送费用
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(UString.parseString(orderVO.getPriceSend(), "0"));
//            nCell.setCellStyle(cellStyle);
//
//
//            //餐盒费
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(UString.parseString(orderVO.getPricePackage(), "0"));
//            nCell.setCellStyle(cellStyle);
//
//
//            //优惠券
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(UString.parseString(orderVO.getPriceCoupon(), "0"));
//            nCell.setCellStyle(cellStyle);
//
//            //站点活动补贴
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(UString.parseString(orderVO.getSiteActivitySubsidy(), "0"));
//            nCell.setCellStyle(cellStyle);
//
//            //小费
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(UString.parseString(orderVO.getTip(), "0"));
//            nCell.setCellStyle(cellStyle);
//
//            //实际支付金额
//            nCell = nRow.createCell(cellNo++);
//            nCell.setCellValue(UString.parseString(orderVO.getPriceActual(), "0"));
//            nCell.setCellStyle(cellStyle);
//
//            //支付方式
//            nCell = nRow.createCell(cellNo++);
//            String method = "";
//            Integer payMethod = orderVO.getPayMethod();
//            if (!UValidator.isNullOrEmpty(payMethod)) {
//                switch (payMethod) {
//                    case 1:
//                        method = "在线支付";
//                        break;
//                    case 2:
//                        method = "货到付款";
//                        break;
//                    default:
//                        break;
//                }
//            }
//            nCell.setCellValue(method);
//            nCell.setCellStyle(cellStyle);
//
//            //配送员
//            nCell = nRow.createCell(cellNo);
//            nCell.setCellValue(orderVO.getSenderName());
//            nCell.setCellStyle(cellStyle);
//
//        }

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
//        wb.write(baos);
        baos.flush();
        baos.close();


//        UPoi.download(baos, response, "订单表.xlsx");
        return null;
    }



}
