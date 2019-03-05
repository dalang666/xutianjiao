package com.atguigu.springboot.utils;



import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import javax.servlet.http.HttpServletResponse;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 导出excel表格数据
 * @ClassName: ExportExcel
 * @author yv.ap666
 * @date 2017年12月12日
 */
public class ExportExcel
{

    private static Class<? extends Object> cls;

    public static <E> void export(HttpServletResponse response, String[] header, String[] fieldNames, String fileName,
                                  List<E> list) throws InvocationTargetException
    {
        HSSFWorkbook wb = new HSSFWorkbook();

        HSSFSheet sheet = wb.createSheet("excel");

        if ((header != null) && (header.length > 0))
        {
            HSSFRow headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++)
            {
                headerRow.createCell(i).setCellValue(header[i]);
            }
        }
        try
        {
            for (int i = 0; i < list.size(); i++)
            {
                HSSFRow contentRow = sheet.createRow(i + 1);

                Object o = list.get(i);
                cls = o.getClass();
                for (int j = 0; j < fieldNames.length; j++)
                {
                    String fieldName = fieldNames[j].substring(0, 1).toUpperCase() + fieldNames[j].substring(1);
                    Method getMethod = cls.getMethod("get" + fieldName, new Class[0]);
                    Object value = getMethod.invoke(o, new Object[0]);
                    if (value != null)
                    {
                        if ((value instanceof Date))
                        {
                            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                            contentRow.createCell(j).setCellValue(sdf.format(value));
                        }
                        else
                        {
                            contentRow.createCell(j).setCellValue(value.toString());
                        }
                    }
                }
            }
        }
        catch (IllegalArgumentException e)
        {
            e.printStackTrace();
        }
        catch (NoSuchMethodException e)
        {
            e.printStackTrace();
        }
        catch (IllegalAccessException e)
        {
            e.printStackTrace();
        }

        OutputStream os = null;
        try
        {
            response.reset();
            response.addHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName,"UTF-8") + ".xls");
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            os = response.getOutputStream();
            wb.write(os);
            os.flush();
        }
        catch (Exception e)
        {
            e.printStackTrace();

            if (os != null) {
                try
                {
                    os.close();
                }
                catch (IOException e1)
                {
                    e1.printStackTrace();
                }
            }
        }
        finally
        {
            if (os != null) {
                try
                {
                    os.close();
                }
                catch (IOException e)
                {
                    e.printStackTrace();
                }
            }
        }
    }
}
