from http import server
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import tkinter.messagebox as mb
import smtplib
import openpyxl

#ダイアログなどに表示されるタイトル
PROGRAM_NAME = "アクセスログメール送信システム"

#対応エクセルファイル名
EXCEL_NAME = 'send_mail_list.xlsx'

#添付ファイルの設置場所
ATTACHMENT_PATH = './attachment/'

#対応エクセルファイルのシート名
SHEET_SETTINGS = 'Send settings'
SHEET_CONTENTS = 'Send contents'

#対応エクセルファイルのセルの位置を指定
#シート1
FROM_ADDRESS_CELL = [2, 3]
SMTP_SERVER_CELL = [3, 3]
SMTP_PORT_CELL = [4, 3]
SMTP_USER_CELL = [5, 3]
SMTP_PASSWORD_CELL = [6, 3]
READ_START_ROW_NO = 9
TO_NAME_COL_NO = 2
TO_ADDRESS_COL_NO = 3
ATTACHMENT_INDIVIDUAL_COL_NO = 4
#シート2
MAIL_SUBJECT_CELL = [2,3]
MAIL_BODY_CELL = [3,3]
ATTACHMENT_COMMON_CELL = [4,3]

#添付ファイルのパスを返す関数
def attachment_path(fail_name):
    return ATTACHMENT_PATH + fail_name

#実行の確認
execution_confirmation = mb.askyesno(
    PROGRAM_NAME,
    'アクセスログのメールを一斉送信します。よろしいですか？'
)

if execution_confirmation:
    #対応エクセルから値の取得
    try:
        book = openpyxl.load_workbook(EXCEL_NAME)
        sheet_send_settings = book[ SHEET_SETTINGS ]
        sheet_send_contents = book[ SHEET_CONTENTS ]
        from_address = sheet_send_settings.cell(*FROM_ADDRESS_CELL).value
        to_address_list = []
        attachment_common = sheet_send_contents.cell(*ATTACHMENT_COMMON_CELL).value
        attachment_individual_list = []

        smtp_host = sheet_send_settings.cell(*SMTP_SERVER_CELL).value
        smtp_port = sheet_send_settings.cell(*SMTP_PORT_CELL).value
        smtp_user = sheet_send_settings.cell(*SMTP_USER_CELL).value
        smtp_pass = sheet_send_settings.cell(*SMTP_PASSWORD_CELL).value

        contents_subject = sheet_send_contents.cell(*MAIL_SUBJECT_CELL).value
        contents_body_before_correction = sheet_send_contents.cell(*MAIL_BODY_CELL).value
        contents_body = contents_body_before_correction.replace('\n', '<br>')

        for row_no in range(READ_START_ROW_NO, sheet_send_settings.max_row+1):
            to_address = sheet_send_settings.cell(row_no, TO_ADDRESS_COL_NO).value
            to_attachment2 = sheet_send_settings.cell(row_no, ATTACHMENT_INDIVIDUAL_COL_NO).value
            if to_address is None:
                break
            to_address_list.append(to_address)
            attachment_individual_list.append(to_attachment2)

        #SMTPサーバーの設定
        server = smtplib.SMTP(smtp_host, smtp_port, timeout=10)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.set_debuglevel(0)
        server.login(smtp_user, smtp_pass)

        #メールの送信
        for i in range(len(to_address_list)):
            message = MIMEMultipart()
            message['Subject'] = contents_subject
            message['From'] = from_address
            message['To'] = to_address_list[i]
            message.attach(MIMEText(contents_body, 'html'))

            attachment_list = [attachment_common, attachment_individual_list[i]]
            for file in attachment_list:
                with open(attachment_path(file), "rb") as f:
                    attachment = MIMEApplication(f.read())
        
                attachment.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=file
                )
                message.attach(attachment)

            server.send_message(message)

        #SMTPサーバーの終了
        server.quit()

        mb.showinfo(PROGRAM_NAME, 'メールの送信が完了しました。')
    except FileNotFoundError:
        mb.showerror(PROGRAM_NAME, f'''必要なファイルが見つかりませんでした。
Excelファイルはこのプログラムと同じディレクトリに{EXCEL_NAME}という名前で設置してください。
添付ファイルは{ATTACHMENT_PATH}というディレクトリをつくりその中に設置してください。''')
else: 
    mb.showinfo(PROGRAM_NAME, '処理は中止されました。')