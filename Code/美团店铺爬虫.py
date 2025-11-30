# API Docs: https://reqable.com/docs/capture/addons

from reqable import *
from datetime import datetime
import json
import re
import os
import win32com.client as win32

def onRequest(context, request):
  # Print url to console
  # print('request url ' + context.url)

  # Update or add a query parameter
  # request.queries['foo'] = 'bar'

  # Update or add a http header
  # request.headers['foo'] = 'bar'

  # Replace http body with a text
  # request.body = 'Hello World'

  # Map with a local file
  # request.body.file('~/Desktop/body.json')

  # Convert to dict if the body is a JSON
  # request.body.jsonify()
  # Update the JSON content
  # request.body['foo'] = 'bar'

  # Done
  return request


def onResponse(context, response):
    # 使用json()方法将响应体解析为字典
    try:
        response.body.jsonify()
    
        # 然后就可以像字典一样访问
        data = response.body["data"]
        baseinfo = data["baseInfo"]
        crumbs = baseinfo["crumbs"]

        # 初始化字典
        all_info_list = []
  
        #提取店铺ID、名称、地址、总体评分、营业时间、WiFi、停车位、店铺分类（简单）、平均价格、最低消费
        shop_id = baseinfo.get("id")
        shop_name = baseinfo.get("name")
        shop_address = baseinfo.get("address")
        shop_score = baseinfo.get("score")
        shop_opentime = baseinfo.get("openTime")
        shop_lowprice = baseinfo.get("lowestPrice")
        
        #平均价格
        if baseinfo["avgPrice"] == 0:
            shop_avgprice = "暂无数据"
        else:
            shop_avgprice = baseinfo.get("avgPrice")
        
        #如果是0，说明无WIFI
        shop_wifi = baseinfo.get("wifi")
        
        #如果是Null，说明无停车位，用0表示
        if baseinfo.get("park") is not None:
            shop_park = 1
        else:
            shop_park = 0
            
        #提取店铺分类
        class_list = []
        for crumb in crumbs:
            category = crumb.get("title", "")
            class_list.append(category)
        shop_class = " ， ".join(class_list)

        # 分别获取经纬度，组合为字符串
        lng = baseinfo.get("lng")
        lat = baseinfo.get("lat")
        shop_location = f"{lng},{lat}"
  
        # 整理数据
        current_data = {
                    "店铺ID": shop_id,
                    "店铺名": shop_name,
                    "店铺分类": shop_class,
                    "评分": shop_score,
                    "人均（元）": shop_avgprice,
                    "最低消费（元）": shop_lowprice,
                    "店铺地址": shop_address,
                    "店铺经纬度": shop_location,
                    "营业时间": shop_opentime,
                    "WIFI": shop_wifi,
                    "停车位": shop_park,
        }
    
        all_info_list.append(current_data)
    
        #导出Excel
        file_path = r"C:\Users\15500\Desktop\店铺评论\店铺评论1.xlsx"
        #先初始化Excel
        exc = None
        workbook = None
        try:
            # 尝试使用Excel.Application
            exc = win32.Dispatch('Excel.Application')
            exc.Visible = False  # 设置Excel窗口不可见（后台运行，不弹出界面）
            if not os.path.exists(file_path):
                workbook = exc.Workbooks.Add()
            else:
                # 打开指定路径的工作簿（Workbook）
                workbook = exc.Workbooks.Open(file_path)
            # 选择活动工作表
            sheet = workbook.ActiveSheet
            # 获取数据起始行
            if sheet.Cells(1, 1).value is None:
                last_row = 1
            else:
                last_row = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row + 1
            print(f'即将从第{last_row}行写入数据')
            # 表头初始化
            if sheet.Cells(1, 1).value is None:
                headers = ["店铺ID", "店铺名", "店铺分类", "评分", "人均（元）", "最低消费（元）", "店铺地址", "店铺经纬度", "营业时间", "WIFI", "停车位"]
                for col, header in enumerate(headers, start=1):
                    sheet.Cells(1, col).value = header
                last_row = 2   # 如果写入表头，从第二行开始写入数据
            #从下一行写入数据主体
            for row_data in all_info_list:
                for col, key in enumerate(["店铺ID", "店铺名", "店铺分类", "评分", "人均（元）", "最低消费（元）", "店铺地址", "店铺经纬度", "营业时间", "WIFI", "停车位"], start=1):
                    value = row_data.get(key)
                    sheet.Cells(last_row, col).value = value
                # 写完一行后，写下一行
                last_row += 1
            
            # 保存工作铺
            if not os.path.exists(file_path):
                workbook.SaveAs(file_path)  # 新建文件需指定路径
            else:
                workbook.Save()  # 已有文件直接保存
                
            print('数据追加成功！')
            print(f"实际保存路径：{os.path.abspath(file_path)}")
    
        # 无论是否出错都执行
        except Exception as e:
            print(f"写入数据失败：{e}")
    
    except Exception as e:
        print(f"处理响应时发生错误: {e}")
        return  # 终止保存流程，避免默认保存
        
    finally:
        # 无论成功与否，都关闭工作簿并退出Excel
        if 'workbook' in locals() and workbook:
            workbook.Close()
        if 'exc' in locals() and exc:
            exc.Quit()
        
    return response


