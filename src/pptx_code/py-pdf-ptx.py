from fpdf import FPDF
import subprocess

pdf = FPDF(orientation="P", format="A4")
pdf.add_page()
pdf.set_margins(0, 0)
pdf.set_font("Arial","B",36)
pdf.set_auto_page_break(auto=False)
a = 2160
b = 3120
pdf.image("Template1.png",w=a/10, h =b/10)
pdf.set_margins(10, 10)
pdf.set_y(80)
pdf.cell(0,h=15,border=0,align="R",txt="Laporan Sentimen")
pdf.set_y(95)
pdf.cell(0,h=15,border=0,align="R",txt="Pemberitaan dan")
pdf.set_y(110)
pdf.cell(0,h=15,border=0,align="R",txt="Media Sosial")

pdf.set_y(150)
pdf.cell(0,h=15,border=0,align="L",txt="Kapolda & Polda")
pdf.set_y(165)
pdf.cell(0,h=15,border=0,align="L",txt="Kapolres & Polres")
pdf.set_font("Arial","B",18)
pdf.set_y(180)
pdf.set_text_color(188,187,187)
pdf.cell(0,h=7,border=0,align="L",txt="03 Agustus - 03 November 2022")

pdf.set_auto_page_break(auto=True)
pdf.set_text_color(0,0,0)
pdf.add_page()
pdf.set_margins(0, 0)
pdf.set_font("Arial","B",16)
pdf.set_auto_page_break(auto=False)
pdf.image("Template3.png",w=a/10, h =b/10)
pdf.set_margins(10, 10)
pdf.set_y(10)
pdf.cell(0,7,border=1,txt="03 Agustus - 03 November 2022",ln=1)
pdf.set_font("Arial","B",36)
pdf.set_y(17)
pdf.cell(0,15,border=1,txt="Polda Aceh")
pdf.set_font("Arial","B",12)
pdf.set_y(50)
pdf.cell(80,15,border=1,txt="Sentimen Pemberitaan Media")
pdf.set_xy(98,52)
pdf.image("Picture5.png")
pdf.set_xy(110,50)
pdf.set_font("Arial","B",24)
pdf.cell(40,15,border=1,txt="1.288")
pdf.set_font("Arial","B",10)
pdf.set_xy(150,55)
pdf.cell(40,10,border=1,txt="Pemberitaan")
pdf.set_xy(10,65)
pdf.cell(90,80,border=1)

data = [{
    "key1":"val1",
    "key2":"val2",
    "key3":"val3",
}]

itx = 100
ity = 65

for e in data:
    for key,val in e.items():
        pdf.set_xy(itx,ity)
        pdf.cell(45,8,border=1)
        itx += 45
        ity += 8
        pdf.set_xy(itx,65)
        pdf.cell(45,8,border=1,ln=1)
        







output = "template_polri.pdf"
pdf.output(output)
# subprocess.run(f"pdf2pptx {output}")



