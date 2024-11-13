import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from exchangelib import Credentials, Account, Message, DELEGATE, Configuration
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
from datetime import datetime
import pandas as pd
import urllib3

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("生日邮件发送系统")
        self.root.geometry("500x400")
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 邮箱输入
        ttk.Label(main_frame, text="发件人邮箱:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.email_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.email_var, width=30).grid(row=0, column=1, columnspan=2, pady=5)
        
        # 密码输入
        ttk.Label(main_frame, text="邮箱密码:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.password_var, show="*", width=30).grid(row=1, column=1, columnspan=2, pady=5)
        
        # 域账号格式说明
        ttk.Label(main_frame, text="格式: hnanet\\username", font=("Arial", 8)).grid(row=2, column=1, sticky=tk.W, pady=2)
        
        # 文件选择
        ttk.Label(main_frame, text="员工信息文件:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.file_path_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.file_path_var, width=30, state='readonly').grid(row=3, column=1, pady=5)
        ttk.Button(main_frame, text="选择文件", command=self.choose_file).grid(row=3, column=2, padx=5)
        
        # 执行按钮
        ttk.Button(main_frame, text="发送生日邮件", command=self.send_birthday_emails).grid(row=4, column=0, columnspan=3, pady=20)
        
        # 状态显示
        self.status_var = tk.StringVar()
        self.status_var.set("等待发送...")
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=5, column=0, columnspan=3, pady=5)
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, length=400, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, pady=5)

    def choose_file(self):
        file_path = filedialog.askopenfilename(
            title="选择员工信息文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)

    def update_status(self, message):
        self.status_var.set(message)
        self.root.update()

    def send_birthday_emails(self):
        # 获取输入的邮箱和密码
        username = self.email_var.get()
        password = self.password_var.get()
        file_path = self.file_path_var.get()
        
        # 验证输入
        if not username or not password:
            messagebox.showerror("错误", "请输入邮箱和密码")
            return
        
        if not file_path:
            messagebox.showerror("错误", "请选择员工信息文件")
            return
        
        self.progress.start()
        self.update_status("正在处理...")
        
        try:
            # 禁用SSL验证
            BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            
            # 发送邮件的主要逻辑
            self.send_birthday_wishes(username, password, file_path)
            
            messagebox.showinfo("成功", "邮件发送完成！")
            self.update_status("发送完成")
            
        except Exception as e:
            messagebox.showerror("错误", f"发送失败：{str(e)}")
            self.update_status("发送失败")
        
        finally:
            self.progress.stop()

    def send_birthday_wishes(self, username, password, file_path):
        data = self.load_employee_data(file_path)
        birthday_employees = self.check_birthdays(data)
        
        if len(birthday_employees) == 0:
            messagebox.showinfo("提示", "今天没有员工过生日")
            return
            
        for _, row in birthday_employees.iterrows():
            name = row["姓名"]
            to_email = row["工作邮箱"]
            self.update_status(f"正在发送给: {name}")
            
            subject = "生日快乐！"
            message = f"亲爱的 {name}, 今天是您的生日，中南维修基地祝您生日快乐！\n温馨提示：公司将为您配送生日蛋糕，如果还没有蛋糕供应商联系您，请你及时与我联系，电话13034983982，谢谢！再次祝您生日快乐！"
            self.send_email(username, password, to_email, subject, message)

    def load_employee_data(self, file_path):
        data = pd.read_excel(file_path)
        data['出生日期'] = pd.to_datetime(data['出生日期'])
        return data

    def check_birthdays(self, data):
        today = datetime.now().strftime("%m-%d")
        birthday_employees = data[data['出生日期'].dt.strftime("%m-%d") == today]
        return birthday_employees

    def send_email(self, username, password, to_email, subject, message):
        try:
            credentials = Credentials(username, password)
            
            config = Configuration(
                server='hybrid.haihangyun.com:443',
                credentials=credentials,
                auth_type='NTLM'
            )
            
            account = Account(
                primary_smtp_address=username if '@' in username else f"{username.split('\\')[1]}@hnair.com",
                config=config,
                autodiscover=False,
                access_type=DELEGATE
            )
            
            m = Message(
                account=account,
                subject=subject,
                body=message,
                to_recipients=[to_email]
            )
            
            m.send()
            
        except Exception as e:
            raise Exception(f"发送给 {to_email} 失败: {str(e)}")

def main():
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
