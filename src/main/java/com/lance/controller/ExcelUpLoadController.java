package com.lance.controller;

import com.lance.domain.DeviceDetails;
import com.lance.utils.CommonUtil;
import com.lance.utils.ExcelRead;
import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author
 * @create 2018-02-03 16:37
 **/
@RestController
public class ExcelUpLoadController {

    private Logger logger = Logger.getLogger(ExcelUpLoadController.class);

    @RequestMapping(value = "/excelUpload", method = {RequestMethod.GET, RequestMethod.POST})
    @ResponseBody
    public Map<String, Object> upload(HttpServletRequest request, HttpServletResponse response, @RequestParam("file") MultipartFile file) {
        logger.info("excel 开始上传");
        //System.out.println("==============excel 开始上传==============");
        Map<String, Object> result = new HashMap<String, Object>();
        if (file.isEmpty()) {
            result.put("result", "1");
            result.put("result", "文件上传失败....");
            result.put("data", null);
        }
        //原文件名
        String name = file.getOriginalFilename();
        //文件大小
        long size = file.getSize();
        if (StringUtils.isBlank(name) && size == 0) {
            result.put("result", "1");
            result.put("result", "未获取到文件");
            result.put("data", null);
        }
        try {
            //读取Excel数据到List中
            List<List<String>> list = new ExcelRead().readExcel(file);
            List<DeviceDetails> deviceDetails = null;
            if (CommonUtil.isExist(list)) {
                deviceDetails = new ArrayList<DeviceDetails>();
                for (List<String> arrayList : list) {
                    if (CommonUtil.isExist(arrayList)) {
                        DeviceDetails device = new DeviceDetails();
                        device.setOpid(arrayList.get(0));
                        device.setOpdt(arrayList.get(1));
                        device.setOptime(arrayList.get(2));
                        device.setD1(arrayList.get(3));
                        device.setD2(arrayList.get(4));

                        device.setD3(arrayList.get(5));
                        device.setD4(arrayList.get(6));
                        device.setD5(arrayList.get(7));
                        device.setD6(arrayList.get(8));
                        device.setD7(arrayList.get(9));

                        device.setD8(arrayList.get(10));
                        device.setD9(arrayList.get(11));
                        device.setD10(arrayList.get(12));
                        device.setD11(arrayList.get(13));
                        device.setD12(arrayList.get(14));

                        device.setD13(arrayList.get(15));
                        device.setD14(arrayList.get(16));
                        device.setD15(arrayList.get(17));

                        deviceDetails.add(device);
                    }
                }
            }

            logger.info("转换后为："+deviceDetails);
            System.out.println("==============end==============");
            result.put("result", "0");
            result.put("result", "上传成功");
            result.put("data", deviceDetails);
        } catch (Exception e) {
            logger.error("读取Excel数据失败", e);
            //System.err.println("==============读取Excel数据失败=============="+e);
        }

        logger.info("excel 上传完成");
        //System.err.println("==============excel 上传完成==============");
        return result;
    }
}
