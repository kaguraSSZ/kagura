import xlrd
import requests
from SMTPSenderEmail import sendEmail


data = xlrd.open_workbook("D:/A_工作/广东医科大学附属医院/病案上传省厅/上传患者用例/副本tmp003.xlsx")  # 读取D盘中excel表格
sheet = data.sheet_by_index(0)  # 取到excel的sheet目录
pid = list(sheet.col_values(1, 1, 231))  # 取sheet的指定列并转换为list类型 第2列 2到231行
vid = list(sheet.col_values(2, 1, 231))  # 第3列  2到231行

url = 'http://10.5.0.41:888/WebService.asmx/UploadOnePat'  # 开始构建POST请求调用本地WebService
headers = {'Content-Type': 'application/x-www-form-urlencoded'}  # 设置POST请求头,欺诈服务器来继续请求
for patientId, visitId in zip(pid, vid):  # 将list中的pid与vid并列显示
    visitId = int(visitId)  # 将visitId转换为int类型
    parameter = {"patientId": patientId, "visitId": visitId}  # 设置参数
    ret = requests.post(url, data=parameter, headers=headers)  # 用requests.post发起post请求
    if ret.content.decode('utf-8').find("false") > -1:
        print("!!!!!患者：%s, 住院次：%s 的患者上传失败!!!!!!" % (str(patientId), str(visitId)))
    else:
        print('----患者id为：%s ,住院次为：%s 的患者已经成功上传----' % (str(patientId), str(visitId)))
        # print(patientId, visitId)

# sendEmail(str(patientId), str(visitId)) 调用发送邮方法发送邮件通知 SMTP协议



