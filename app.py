import time
from flask import Flask, render_template
import webbrowser
from threading import Timer
from flask import request
import xlwt
from xlwt import Workbook
from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import xlwings as xw
import datetime
import os 
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import threading




app = Flask(__name__, template_folder='templates')


@app.route('/home', methods=["GET", "POST"])
def home():
    return render_template("home.html")


def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000/home")


def send_email(sender_email, sender_password, recipient_emails, subject, body, pdf_file_path, id):
    # Create message container
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipient_emails)  # Join recipient emails with commas
    msg['Subject'] = subject

    # Attach message body
    msg.attach(MIMEText(body, 'plain'))

    # Attach PDF file
    try:
        with open(pdf_file_path, 'rb') as pdf_file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(pdf_file.read())
            encoders.encode_base64(part)
            pdf_fillename = f"receipt{id}.pdf"
            part.add_header('Content-Disposition', f'attachment; filename={pdf_fillename}')
            msg.attach(part)
    except Exception as e:
        print(f"Failed to attach file: {e}")
        return

    try:
        # Establish connection to Gmail server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()  # Secure the connection
        server.login(sender_email, sender_password)

        # Send the email
        server.sendmail(sender_email, recipient_emails, msg.as_string())
        server.quit()

        print("Email sent successfully! to " + ", ".join(recipient_emails))
    except Exception as e:
        print(f"Failed to send email: {e}")


@app.route('/end', methods=["GET", "POST"])
def end():

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    if "ceramic" in request.form.get("select1"):
        ws = xw.Book(r'טבלאות\ceramic table.xls').sheets['Sheet 1']
    elif "speech" in request.form.get("select1"):
        ws = xw.Book(r'טבלאות\speech therapy table.xls').sheets['Sheet 1']

    count = 2
    while ws.range("A" + str(count + 1)).value != None:
        count += 1
    id = int(ws.range("A" + str(count)).value) + 1
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase import pdfmetrics

    pdfmetrics.registerFont(TTFont('Hebrew', 'C:/Windows/Fonts/David.ttf'))
    can.setFont('Hebrew', 12)
    can.drawString(200, 543, str(id))
    ws.range("A" + str(count + 1)).value = int(ws.range("A" + str(count)).value) + 1
    count += 1

    name = request.form.get("name")
    width = 470
    # for lett in name:
    #     if lett == " ":
    #         width -= 5.5
    #     else:
    #         width -= 7

    can.drawString(width - can.stringWidth(name[::-1], 'Hebrew', 12), 518, name[::-1])
    #can.drawString(170, 518, name)

    i = 0
    highet = 460
    prices = []
    item = request.form.get("item1")
    can.drawString(490 - can.stringWidth(item[::-1], 'Hebrew', 12) , highet, item[::-1])
    prices += [int(request.form.get("price1"))]
    can.drawString(180, highet, str(prices[i]))
    i += 1
    while request.form.get("item"+str(i+1)):
        highet -= 25
        item = request.form.get("item"+str(i+1))
        can.drawString(490 - can.stringWidth(item[::-1], 'Hebrew', 12) , highet, item[::-1])
        prices += [int(request.form.get("price"+str(i+1)))]
        can.drawString(180, highet, str(prices[i]))
        i += 1

    can.drawString(130, 258, str(sum(prices)))
    ws.range("C" + str(count)).value = sum(prices)
    date = request.form.get("date", datetime.datetime.now().strftime("%Y-%m-%d"))
    try:
        import locale
        locale.setlocale(locale.LC_TIME, 'en_US.UTF-8')
        date_obj = datetime.datetime.strptime(date, "%Y-%m-%d")
        date = date_obj.strftime("%d/%m/%Y")
    except ValueError:
        date = datetime.datetime.now().strftime("%d/%m/%Y")
    can.drawString(355, 173, date)
    ws.range("B" + str(count)).value = date
    # Save the Excel table changes
    ws.book.save()
    ws.book.close()
    os.system("taskkill /f /im excel.exe")

    can.save()

    # move to the beginning of the StringIO buffer
    packet.seek(0)

    # create a new PDF with Reportlab
    new_pdf = PdfFileReader(packet)  # PdfFileReader
    # read your existing PDF
    if "ceramic" in request.form.get("select1"):
        existing_pdf = PdfFileReader(open(r"more files\generic2.pdf", "rb"))  # PdfFileReader
    elif "speech" in request.form.get("select1"):
        existing_pdf = PdfFileReader(open(r"more files\generic3.pdf", "rb"))  # PdfFileReader
    output = PdfFileWriter()  # PdfFileWriter
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    l = "ds" + "ds"
    # finally, write "output" to a real file
    pdf_file_path = ""
    if "ceramic" in request.form.get("select1"):
        pdf_file_path = r"קבלות קרמיקה" + r"\קבלה " + str(id) + " " + str(name) + ".pdf"
        outputStream = open(pdf_file_path, "wb")  # C:\Users\IMOE001\Desktop\אמא\קבלה ירוקה\קבלות עבר
    elif "speech" in request.form.get("select1"):
        pdf_file_path = r"קבלות קלינאות תקשורת" + r"\קבלה " + str(id) + " " + str(name) + ".pdf"
        outputStream = open(pdf_file_path, "wb")  # C:\Users\IMOE001\Desktop\אמא\קבלה ירוקה\קבלות עבר

    # outputStream = open(r"C:\Users\tzurd\OneDrive\Desktop\אמא\קבלה ירוקה\קבלות עבר" + r"\קבלה " + str(id) + " " + str(name) + ".pdf", "wb")
    output.write(outputStream)
    outputStream.close()

    email = request.form.get("email")
    if email and email != "":
        # dir_path = os.path.dirname(os.path.realpath(__file__))
        # pdf_file_path = os.path.join(dir_path, pdf_file_path) # absulute path
        if "ceramic" in request.form.get("select1"):
            subject = "אפרת דנינו - קרמיקה - קבלה מספר " + str(id)
        elif "speech" in request.form.get("select1"):
            subject ="אפרת דנינו - קלינאות תקשורת - קבלה מספר " + str(id)
        body = "הקבלה מצורפת"
        emails = [email]
        sendToSelf = request.form.get("sendToSelf")
        if sendToSelf == "on":
            emails.append('efratda1978@gmail.com')
        send_email('receiptsgreen33@gmail.com', 'sojw dccy jexf vbgu', emails, subject, body, pdf_file_path, id)
    threading.Timer(10, lambda: os._exit(0)).start()
    return render_template("end.html")


if __name__ == '__main__':
    Timer(1, open_browser).start()
    app.run(host='0.0.0.0', port=5000)
