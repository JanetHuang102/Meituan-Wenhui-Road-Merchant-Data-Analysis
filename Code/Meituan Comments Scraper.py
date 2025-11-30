# API Docs: https://reqable.com/docs/capture/addons

from reqable import *
from datetime import datetime
import json
import re
import win32com.client as win32

# 时间戳转换函数
def timestamp_to_datetime(ms_timestamp):
    try:
        sec_timestamp = ms_timestamp / 1000
        dt = datetime.fromtimestamp(sec_timestamp)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return "未知时间"


# 合并评论的函数
def merge_with_merchant_comment(original_content, merchant_comment):
    merged = [f"原评论：{original_content}"]
    if merchant_comment is not None:
        reply_content = merchant_comment
        merged.append(f"【商家回复】：{reply_content}")
    return "\n".join(merged)


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
    # 将相应体转换为JSON对象
    response.body.jsonify()
    origin_commentlist = response.body["comments"]
  
    # 初始化字典
    all_comment_list = []
  
    #提取用户ID、评论时间、用户评分、评论具体内容
    for comment in origin_commentlist:
        user_name = comment.get("userName")
        modify_time = comment.get('modifyTime', 0)
        comment_time = timestamp_to_datetime(modify_time)
        star = comment.get("star", 0) / 10
    
        # 如果有商家回复、其他用户回复，则和用户评论合并为一条评论
        merchant_comment = comment.get("merchantComment")
        comment_body = comment.get("commentBody")
        merge_comment = merge_with_merchant_comment(comment_body, merchant_comment)
  
        # 整理数据
        current_data = {
                    "用户名": user_name,
                    "评分": star,
                    "评论内容": merge_comment,
                    "评论时间": comment_time
        }
    
        all_comment_list.append(current_data)
    
    # 字符转换
    all_comment_str = json.dumps(all_comment_list, ensure_ascii=False)
    
    # json列表转回python
    match = re.search(r'\[.*\]', all_comment_str)
    if match:
        all_comment_list = eval(match.group())
    else:
        all_comment_list = []
    
    #导出Excel
    file_path = r"C:\Users\15500\Desktop\美团评论\美团评论.xlsx"
  
    # 检查文件夹是否存在，不存在则创建
    dir_path = os.path.dirname(file_path)
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)

    try:
        # 尝试使用Excel.Application
        exc = win32.Dispatch('Excel.Application')
        exc.Visible = False  # 设置Excel窗口不可见（后台运行，不弹出界面）
        # 检查文件是否存在，不存在则创建新工作簿
        if os.path.exists(file_path):
            workbook = exc.Workbooks.Open(file_path)
        else:
            workbook = exc.Workbooks.Add()  # 创建新工作簿
            workbook.SaveAs(file_path)  # 保存到目标路径
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
            headers = ['用户名', '评分', '评论内容', '评论时间']
            for col, header in enumerate(headers, start=1):
                sheet.Cells(1, col).value = header
            last_row = 2   # 如果写入表头，从第二行开始写入数据
        #从下一行写入数据主体
        for row_data in all_comment_list:
            for col, key in enumerate(['用户名', '评分', '评论内容', '评论时间'], start=1):
                value = row_data.get(key)
                sheet.Cells(last_row, col).Value = value
            # 写完一行后，写下一行
            last_row += 1
        # 保存工作铺
        workbook.Save()
        print('数据追加成功！')
    
    # 无论是否出错都执行
    except Exception as e:
        print(f"写入数据失败：{e}")
    
    finally:
        # 无论成功与否，都关闭工作簿并退出Excel
        if 'workbook' in locals() and workbook:
            workbook.Close()
        if 'exc' in locals() and exc:
            exc.Quit()
    
    return response

