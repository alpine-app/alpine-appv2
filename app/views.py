"""
Definition of views.
"""

from datetime import datetime, date
from django.shortcuts import render
from django.conf import settings
from django.http import HttpRequest, HttpResponseRedirect, HttpResponse
from app import models
from app import forms
from django.contrib import messages
from django.core.files.storage import FileSystemStorage
import random, os, io
import pandas as pd
from docx import Document

#Các hàm xử lý văn bản cho ngân hàng câu hỏi đưa lên trước
def get_run_text_with_formatting(run):
    text = run.text 
    if not text:
        return ""
        
    if run.bold:
        text = f"<b>{text}</b>"
    if run.italic:
        text = f"<i>{text}</i>"
    if run.underline:
        text = f"<u>{text}</u>"
        
    return text

def extract_docx_with_formatting(file_path):
    mdoc = Document(file_path)
    full_text = []

    for element in mdoc.element.body:
        # Xử lý các đoạn văn bản (Paragraphs)
        if element.tag.endswith('p'):
            para_text = []
            paragraph = mdoc.paragraphs[len([e for e in mdoc.element.body[:mdoc.element.body.index(element)] if e.tag.endswith('p')])]
            
            # Xử lý tiêu đề (Headings)
            if paragraph.style.name.startswith('Heading'):
                level = paragraph.style.name.split()[-1]
                para_text.append(f"<h{level}>")
                
            for run in paragraph.runs:
                para_text.append(get_run_text_with_formatting(run))
                
            if paragraph.style.name.startswith('Heading'):
                para_text.append(f"</h{level}>")
                
            full_text.append(" " + "".join(para_text))

        # Xử lý bảng (Tables)
        elif element.tag.endswith('tbl'):
            table_idx = len([e for e in mdoc.element.body[:mdoc.element.body.index(element)] if e.tag.endswith('tbl')])
            table = mdoc.tables[table_idx]
            
            table_html = ["<table border='1'>"]
            for row in table.rows:
                table_html.append("<tr>")
                for cell in row.cells:
                    table_html.append("<td>")
                    for para in cell.paragraphs:
                        for run in para.runs:
                            table_html.append(get_run_text_with_formatting(run))
                    table_html.append("</td>")
                table_html.append("</tr>")
            table_html.append("</table>")
            full_text.append("".join(table_html))

    return full_text

def khop_ABCD(text):
    #tìm khớp ký tự A. B. C. D.
    chuoih = ["A.", "B.", "C.", "D.", "E.", "F.", "G.", "H."]
    chuoi = ["a.", "b.", "c.", "d.", "e.", "f.", "g.", "h."]
    results = []
    # Tách chuỗi thành các từ/cụm từ dựa trên khoảng trắng (nếu text chứa nhiều câu)
    # Hoặc xử lý trực tiếp nếu text là một chuỗi liên tục chứa đáp án
    words = text.split()
    for word in words:
        # Loại bỏ các thẻ HTML bằng cách duyệt từng ký tự
        clean_word = ""
        inside_tag = False
        for char in word:
            if char == '<':
                inside_tag = True
            elif char == '>':
                inside_tag = False
            elif not inside_tag:
                clean_word += char
        # Kiểm tra xem chuỗi sạch có phải là "A.", "B.", "C.", "D." hay không
        for i in range(8) :
            if clean_word == chuoih[i] or clean_word == chuoi[i]:
                results.append([chuoih[i]])
                break
        #if clean_word in ["A.", "B.", "C.", "D.", "E.", "F.", "G.", "H."]:
        #    results.append(clean_word)
    return results

def khop_ans(text):
    #tìm khớp ký tự <u>A.</u>
    chuoih = ["A.", "B.", "C.", "D.", "E.", "F.", "G.", "H."]
    chuoi = ["a.", "b.", "c.", "d.", "e.", "f.", "g.", "h."]
    results = []
    # Tách chuỗi thành các từ/cụm từ dựa trên khoảng trắng (nếu text chứa nhiều câu)
    # Hoặc xử lý trực tiếp nếu text là một chuỗi liên tục chứa đáp án
    words = text.split()
    for word in words:
        # Loại bỏ các thẻ HTML bằng cách duyệt từng ký tự
        clean_word = ""
        inside_tag = False
        co_u1 = 0
        co_u = False
        for char in word:
            if char == '<':
                inside_tag = True
                co_u1 = 1
            elif char == '>':
                inside_tag = False
                co_u1 = 0
            elif not inside_tag:
                clean_word += char
            elif char == 'u' and co_u1 == 1:
                co_u = True
                co_u1 = 0
            elif char == '/' and co_u1 == 1:
                co_u1 = 0
        # Kiểm tra xem chuỗi sạch có phải là "A.", "B.", "C.", "D." hay không
        if co_u :
            for i in range(8) :
                if clean_word == chuoih[i] or clean_word == chuoi[i]:
                    results.append(chuoih[i])
                    co_u = False
                    break
        #if clean_word in ["A.", "B.", "C.", "D.", "E.", "F.", "G.", "H."] and co_u:
        #    results.append(clean_word)
        #    co_u = False
    return results

def up_contents_lc(conml_contentx, mmon, sott, mdvkt, tendvkt, mtd, mdo):
    #tao bieen loi
    loi = ""
    #lisst so sanh
    lssanh = [["A."], ["B."], ["C."], ["D."], ["E."], ["F."], ["G."], ["H."]]
    tt = sott
    mch = mmon + f"{tt:04d}"
    datenum = date.today()
    datesua = datenum.strftime('%Y-%m-%d')
    sm_hcontents = [{"macauhoi": mch, "titles": mtd, "questions": "Noi dung cau hoi", "picques": "", "audques": "", "clipques": "", "permuteques": True,
            "answer": "Dap an 1<##>Dap an 2<##>Dap an 3", "picans": "", "audans": "", "clipans": "", "permuteans": "T",
            "classifyques": "LC, DT, DS, NH, NL, TL", "partques": tendvkt, "levelques": mdo, "correctans": "", "layoutans1": 1, "mamon": mmon, "madvkt": mdvkt, "createdate": datesua}]
    htongdoan = len(conml_contentx)
    if htongdoan > 1:
        inchayfirst = 0
        inchay = inchayfirst
        chay = 0
        if "[<S>]" in conml_contentx[chay] or "[<s>]" in conml_contentx[chay]:
            sm_hcontents[inchay]["classifyques"] = "LC"
            if "[<S>]" in conml_contentx[chay]:
                vt = conml_contentx[chay].find("[<S>]") + 5
            else:
                vt = conml_contentx[chay].find("[<s>]") + 5
            sm_hcontents[inchay]["titles"] = mtd
            chay += 1
            while chay < htongdoan - 1:
                sm_hcontents[inchay]["questions"] = conml_contentx[chay]
                chay += 1
                s = conml_contentx[chay]
                if "[<br>]" in s or "[<BR>]" in s:
                    break
                while chay < htongdoan:
                    khop = khop_ABCD(s)
                    if khop != [] :
                        break
                    sm_hcontents[inchay]["questions"] = sm_hcontents[inchay]["questions"] + "<br>" + s
                    chay += 1
                    if chay < htongdoan:
                        s = conml_contentx[chay]
                        if "[<br>]" in s or "[<BR>]" in s:
                            loi = "Khong co dap an"
                            break
                    else:
                        break
                if chay < htongdoan:
                    s = conml_contentx[chay]
                    if "[<br>]" in s or "[<BR>]" in s:
                        loi = "Khong co dap an"
                    else:
                        khop = khop_ABCD(s)
                        if len(khop) >= 3:
                            #Nếu cùng 1 hàng thì
                            khopans = khop_ans(s)
                            if khopans == ["A."]:
                                sm_hcontents[inchay]["correctans"] = "A"
                            elif khopans == ["B."]:
                                sm_hcontents[inchay]["correctans"] = "B"
                            elif khopans == ["C."]:
                                sm_hcontents[inchay]["correctans"] = "C"
                            elif khopans == ["D."]:
                                sm_hcontents[inchay]["correctans"] = "D"
                            elif khopans == ["E."]:
                                sm_hcontents[inchay]["correctans"] = "E"
                            elif khopans == ["F."]:
                                sm_hcontents[inchay]["correctans"] = "F"
                            elif khopans == ["G."]:
                                sm_hcontents[inchay]["correctans"] = "G"
                            else:
                                sm_hcontents[inchay]["correctans"] = "A"
                            trich1 = s.replace(r"<u>", "")
                            s1 = trich1.replace(r"</u>", "")
                            s1 = s1.replace("\t", " ")
                            vta = s1.find("A. ") + 2
                            trich1 = s1[vta:]
                            vtb = trich1.find("B. ")
                            #print(trich1)
                            sm_hcontents[inchay]["answer"] = trich1[:vtb]
                            sm_hcontents[inchay]["permuteans"] = "T"
                            vtb = vtb + 2
                            trich1 = trich1[vtb:]
                            #-------------------
                            vtb = trich1.find("C. ")
                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1[:vtb]
                            sm_hcontents[inchay]["permuteans"] = sm_hcontents[inchay]["permuteans"] + ",T"
                            vtb = vtb + 2
                            trich1 = trich1[vtb:]
                            #-------------------
                            vtb = trich1.find("D. ")
                            if vtb == -1:
                                sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1
                            else:
                                sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1[:vtb]
                                vtb = vtb + 3
                                trich1 = trich1[vtb:]
                                sm_hcontents[inchay]["permuteans"] = sm_hcontents[inchay]["permuteans"] + ",T"
                            #-------------------
                            vtb = trich1.find("E. ")
                            if vtb == -1:
                                sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1
                            else:
                                sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1[:vtb]
                                vtb = vtb + 3
                                trich1 = trich1[vtb:]
                                sm_hcontents[inchay]["permuteans"] = sm_hcontents[inchay]["permuteans"] + ",T"
                            chay += 1
                            #print(chay)
                        elif len(khop) == 2:
                            khopans = khop_ans(s)
                            if khopans == ["A."]:
                                sm_hcontents[inchay]["correctans"] = "A"
                            elif khopans == ["B."]:
                                sm_hcontents[inchay]["correctans"] = "B"
                            
                            trich1 = s.replace(r"<u>", "")
                            s1 = trich1.replace(r"</u>", "")
                            s1 = s1.replace(r"\t", " ")
                            vta = s.find("A. ") + 2
                            trich1 = s[vta:]
                            vtb = trich1.find("B. ")
                            sm_hcontents[inchay]["answer"] = trich1[:vtb]
                            vtb = vtb + 2
                            trich1 = trich1[vtb:]
                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1
                            #-Xu ly tiep các đáp án sau ở hàng kế tiếp
                            chay += 1
                            if chay < htongdoan and "[<br>]" not in conml_contentx[chay]:
                                s = conml_contentx[chay]
                                khop = khop_ABCD(s)
                                if all(item in lssanh for item in khop):
                                    khopans = khop_ans(s)
                                    if khopans == ["C."] :
                                        sm_hcontents[inchay]["correctans"] = "C"
                                    elif khopans == ["D."] :
                                        sm_hcontents[inchay]["correctans"] = "D"
                                    trich1 = s.replace(r"<u>", "")
                                    s1 = trich1.replace(r"</u>", "")
                                    s1 = s1.replace("\t", " ")
                                    vta = s1.find("C. ") + 2
                                    vtb = s1.find("D. ")
                                    #print("kiểm tra chỗ này")
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s1[vta:vtb]
                                    vtb = vtb + 2
                                    trich1 = s1[vtb:]
                                    #-------------------
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1
                                    chay += 1
                        else:
                            khopans = khop_ans(s)
                            if khopans == ["A."] :
                                sm_hcontents[inchay]["correctans"] = "A"
                                s = conml_contentx[chay]
                                vt = s.find(r"/u") + 3
                                sm_hcontents[inchay]["answer"] = s[vt:]
                                chay += 1
                                while (chay < htongdoan and "[<br>]" not in conml_contentx[chay]):
                                    khop1 = khop_ABCD(conml_contentx[chay])
                                    if all(item in lssanh for item in khop1):
                                        break
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + conml_contentx[chay]
                                    chay += 1
                            else:
                                s = conml_contentx[chay]
                                vt = s.find(".") + 2
                                sm_hcontents[inchay]["answer"] = s[vt:]
                                chay += 1
                                while (chay < htongdoan and "[<br>]" not in conml_contentx[chay]):
                                    khop1 = khop_ABCD(conml_contentx[chay])
                                    if all(item in lssanh for item in khop1):
                                        break
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + conml_contentx[chay]
                                    chay += 1
                            sm_hcontents[inchay]["permuteans"] = "T"
                            #Thụt vào 
                            if chay < htongdoan:
                                s = conml_contentx[chay]
                                if "[<br>]" in s or "[<BR>]" in s:
                                    #print('Khong co dap an đây')
                                    loi = "Khong co dap an đây"
                                else:
                                    khopans = khop_ans(s)
                                    if khopans == ["B."] :
                                        if sm_hcontents[inchay]["correctans"] == "":
                                            sm_hcontents[inchay]["correctans"] = "B"
                                        else:
                                            #print('Da co dap an A - trung dap an')
                                            loi = "Da co dap an A - trung dap an"
                                        s = conml_contentx[chay]
                                        vt = s.find(r"/u") + 3
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in conml_contentx[chay]):
                                            khop1 = khop_ABCD(conml_contentx[chay])
                                            if all(item in lssanh for item in khop1):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + conml_contentx[chay]
                                            chay += 1
                                    else:
                                        s = conml_contentx[chay]
                                        vt = s.find(".") + 2
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in conml_contentx[chay]):
                                            khop1 = khop_ABCD(conml_contentx[chay])
                                            if all(item in lssanh for item in khop1):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + conml_contentx[chay]
                                            chay += 1
                                    sm_hcontents[inchay]["permuteans"] = sm_hcontents[inchay]["permuteans"] + ",T"
                            if chay < htongdoan:
                                s = conml_contentx[chay]
                                if "[<br>]" not in s and "[<BR>]" not in s:
                                    khopans = khop_ans(s)
                                    if khopans == ["C."] :
                                        if sm_hcontents[inchay]["correctans"] == "":
                                            sm_hcontents[inchay]["correctans"] = "C"
                                        else:
                                            #print('Da co dap an A, B - trung dap an')
                                            loi = "Da co dap an A, B - trung dap an"
                                        s = conml_contentx[chay]
                                        vt = s.find(r"/u") + 3
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in conml_contentx[chay]):
                                            khop1 = khop_ABCD(conml_contentx[chay])
                                            if all(item in lssanh for item in khop1):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + conml_contentx[chay]
                                            chay += 1
                                    else:
                                        s = conml_contentx[chay]
                                        vt = s.find(".") + 2
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in conml_contentx[chay]):
                                            khop1 = khop_ABCD(conml_contentx[chay])
                                            if all(item in lssanh for item in khop1):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + conml_contentx[chay]
                                            chay += 1
                                    sm_hcontents[inchay]["permuteans"] = sm_hcontents[inchay]["permuteans"] + ",T"
                            if chay < htongdoan:
                                s = conml_contentx[chay]
                                if "[<br>]" not in s and "[<BR>]" not in s:
                                    khopans = khop_ans(s)
                                    if khopans == ["D."] :
                                        if sm_hcontents[inchay]["correctans"] == "":
                                            sm_hcontents[inchay]["correctans"] = "D"
                                        else:
                                            #print('Da co dap an A, B, C - trung dap an')
                                            loi = "Da co dap an A, B, C - trung dap an"
                                        s = conml_contentx[chay]
                                        vt = s.find(r"/u") + 3
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in conml_contentx[chay]):
                                            khop1 = khop_ABCD(conml_contentx[chay])
                                            if all(item in lssanh for item in khop1):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + conml_contentx[chay]
                                            chay += 1
                                    else:
                                        s = conml_contentx[chay]
                                        vt = s.find(".") + 2
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in conml_contentx[chay]):
                                            khop1 = khop_ABCD(conml_contentx[chay])
                                            if all(item in lssanh for item in khop1):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + conml_contentx[chay]
                                            chay += 1
                                    sm_hcontents[inchay]["permuteans"] = sm_hcontents[inchay]["permuteans"] + ",T"
                            if chay < htongdoan:
                                s = conml_contentx[chay]
                                if "[<br>]" not in s and "[<BR>]" not in s:
                                    khopans = khop_ans(s)
                                    if khopans == ["E."] :
                                        if sm_hcontents[inchay]["correctans"] == "":
                                            sm_hcontents[inchay]["correctans"] = "E"
                                        else:
                                            #print('Da co dap an A, B, C, D - trung dap an')
                                            loi = "Da co dap an A, B, C, D - trung dap an"
                                        s = conml_contentx[chay]
                                        vt = s.find(r"/u") + 3
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in conml_contentx[chay]):
                                            khop1 = khop_ABCD(conml_contentx[chay])
                                            if all(item in lssanh for item in khop1):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + conml_contentx[chay]
                                            chay += 1
                                    else:
                                        s = conml_contentx[chay]
                                        vt = s.find(".") + 2
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in conml_contentx[chay]):
                                            khop1 = khop_ABCD(conml_contentx[chay])
                                            if all(item in lssanh for item in khop1):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + conml_contentx[chay]
                                            chay += 1
                                    sm_hcontents[inchay]["permuteans"] = sm_hcontents[inchay]["permuteans"] + ",T"
                if (chay < htongdoan and "[<br>]" in conml_contentx[chay]) or (chay < htongdoan and "[<BR>]" in conml_contentx[chay]):
                    chay += 1
                    tt = tt + 1
                    mch = mmon + f"{tt:04d}"
                    datenum = date.today()
                    datesua = datenum.strftime('%Y-%m-%d')
                    sm_hcontents.append({"macauhoi": mch, "titles": mtd, "questions": "Noi dung cau hoi", "picques": "", "audques": "", "clipques": "", "permuteques": True,
                        "answer": "Dap an 1<##>Dap an 2<##>Dap an 3", "picans": "", "audans": "", "clipans": "", "permuteans": "T",
                        "classifyques": "LC", "partques": tendvkt, "levelques": mdo, "correctans": "", "layoutans1": 1, "mamon": mmon, "madvkt": mdvkt, "createdate": datesua})
                
                    inchay += 1
                    sm_hcontents[inchay]["classifyques"] = "LC"
                    sm_hcontents[inchay]["titles"] = mtd
    else:
        loi = "Tài liệu không có đoạn văn bản nào."
    return sm_hcontents, loi

def up_contents_nh(chtml_content, mmon, sott, mdvkt, tendvkt, mtd, mdo):
    #tao bieen loi
    loi = ""
    #lisst so sanh
    lssanh = [["A."], ["B."], ["C."], ["D."], ["E."], ["F."], ["G."], ["H."]]
    tt = sott
    mch = mmon + f"{tt:04d}"
    datenum = date.today()
    datesua = datenum.strftime('%Y-%m-%d')
    sm_hcontents = [{"macauhoi": mch, "titles": mtd, "questions": "Noi dung cau hoi", "picques": "", "audques": "", "clipques": "", "permuteques": False,
            "answer": "Dap an 1<##>Dap an 2<##>Dap an 3", "picans": "", "audans": "", "clipans": "", "permuteans": "T",
            "classifyques": "LC, DT, DS, NH, NL, TL", "partques": tendvkt, "levelques": mdo, "correctans": "", "layoutans1": 1, "mamon": mmon, "madvkt": mdvkt, "createdate": datesua}]
    htongdoan = len(chtml_content)
    if htongdoan > 1:
        inchayfirst = 0
        inchay = inchayfirst
        chay = 0
        if "[<G>]" in chtml_content[chay] or "[<g>]" in chtml_content[chay]:
        #print("phan viet cho cau hoi nhom chung")
            sm_hcontents[inchay]["classifyques"] = "NH"
            if "[<G>]" in chtml_content[chay]:
                vt = chtml_content[chay].find("[<G>]") + 5
            else:
                vt = chtml_content[chay].find("[<g>]") + 5
            #sm_hcontents[inchay]["titles"] = chtml_content[chay][vt:]
            sm_hcontents[inchay]["titles"] = mtd
            chay += 1
            # Lay noi dung tung doan nhom
            while chay < htongdoan - 2:
                sm_hcontents[inchay]["questions"] = chtml_content[chay]
                chay += 1
                s = chtml_content[chay]
                if (chay >= htongdoan - 1) and ("[<br>]" in s or "[<BR>]" in s):
                    break
                while chay < htongdoan - 1 and "(<1>)" not in s:
                    sm_hcontents[inchay]["questions"] = sm_hcontents[inchay]["questions"] + "<br>" + s
                    chay += 1
                    if chay < htongdoan:
                        s = chtml_content[chay]
                        if "[<br>]" in s or "[<BR>]" in s:
                            #print('Khong co phần câu hỏi của nhóm')
                            loi = "Khong co phần câu hỏi của nhóm"
                            break
                    else:
                        break
                # Phần tìm câu hỏi và đáp án cho nhóm
                h_cau = 1
                # đặt trc answer ="" để khởi tạo và luôn nối vào trong toàn bộ phần nhóm chung
                sm_hcontents[inchay]["answer"] = ""
                # khởi tạo không xáo trộn các câu
                sm_hcontents[inchay]["permuteans"] = "T"
                while chay < htongdoan and f"(<{h_cau}>)" in chtml_content[chay]:
                    # Lấy nội dung câu hỏi 1
                    # Thêm vào để xử lý các đáp án nằm cùng dòng với câu hỏi (câu hỏi trống)
                    s = chtml_content[chay]
                    ass_cauhoi = ""
                    h_dapan = ""
                    khop = khop_ABCD(s)
                    #aok = ("\tA. " in s or "A. " in s or "a. " in s or "\t<u>A.</u>" in s or r"<u>A.</u>" in s or r"<u>a.</u>" in s or r"<u>A.</u><u>.</u>" in s or r"<u>a.</u><u>.</u>" in s)
                    #bok = ("\tB. " in s or "B. " in s or "b. " in s or "\t<u>B.</u>" in s or r"<u>B.</u>" in s or r"<u>b.</u>" in s or r"<u>B.</u><u>.</u>" in s or r"<u>b.</u><u>.</u>" in s)
                    #cok = ("\tC. " in s or "C. " in s or "c. " in s or "\t<u>C.</u>" in s or "<u>C.</u>" in s or r"<u>c.</u>" in s or r"<u>C.</u><u>.</u>" in s or r"<u>c.</u><u>.</u>" in s)
                    #dok = ("\tD. " in s or "D. " in s or "d. " in s or "\t<u>D.</u>" in s or "<u>D.</u>" in s or r"<u>d.</u>" in s or r"<u>D.</u><u>.</u>" in s or r"<u>d.</u><u>.</u>" in s)
                    if len(khop) >= 3 :
                        # xu ly cau va cac dap an cung 1 hang
                        khopans = khop_ans(s)
                        #if r"<u>A.</u>" in s or r"<u>a.</u>" in s or r"<u>A.</u><u>.</u>" in s or r"<u>a.</u><u>.</u>" in s:
                        if khopans != [] :
                            h_dapan = f"{h_cau}" + khopans[0][:1]
                        else:
                            h_dapan = f"{h_cau}A"
                        #print("dap an 1 la " + h_dapan)
                        trich1 = s.replace("<u>", "")
                        s1 = trich1.replace(r"</u>", "")
                        s1 = s1.replace("\t", " ")
                        vta = s1.find("A. ")
                        v0 = s1.find(f"{h_cau}>)") + 3#lấy vị trí đầu của câu hỏi con
                        #Nếu vta >=0 thì lấy câu hỏi con trước các đáp án abcd và bao trong cặp [< >]
                        if vta>=0:
                            ass_cauhoi = "[<" + s1[v0:vta] + ">]"
                            #sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "[<" + s1[v0:vta] + ">]"
                        #--------------------
                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + ass_cauhoi
                        vta = vta + 3
                        trich1 = s1[vta:]
                        vtb = trich1.find("B. ")
                        #print(trich1)
                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + trich1[:vtb]
                        vtb = vtb + 3
                        trich1 = trich1[vtb:]
                        #-------------------
                        vtb = trich1.find("C. ")
                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1[:vtb]
                        vtb = vtb + 3
                        trich1 = trich1[vtb:]
                        #-------------------
                        vtb = trich1.find("D. ")
                        if vtb == -1:
                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1
                        else:
                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1[:vtb]
                            vtb = vtb + 3
                            trich1 = trich1[vtb:]
                        #-------------------
                        vtb = trich1.find("E. ")
                        if vtb == -1:
                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1
                        else:
                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1[:vtb]
                            vtb = vtb + 3
                            trich1 = trich1[vtb:]
                        chay += 1
                    elif len(khop) == 2 :
                        # xu ly cau va 2 dap an cung 1 hang
                        khopans = khop_ans(s)
                        if khopans == ["A."] :
                        #if r"<u>A.</u>" in s or r"<u>a.</u>" in s or r"<u>A.*</u>" in s or r"<u>a.*</u>" in s:
                            h_dapan = f"{h_cau}A"
                        elif khopans == ["B."] :
                            h_dapan = f"{h_cau}B"
                            
                        trich1 = s.replace("<u>", "")
                        s1 = trich1.replace(r"</u>", "")
                        vta = s1.find("A. ")
                        v0 = s1.find(f"{h_cau}>)") + 3#lấy vị trí đầu của câu hỏi con
                        #Nếu vta >=0 thì lấy câu hỏi con trước các đáp án abcd và bao trong cặp [< ]<
                        if vta>=0:
                            sm_hcontents[inchay]["answer"] = "[<" + s1[v0:vta] + ">]"
                        #--------------------
                        vta = vta + 3
                        trich1 = s[vta:]
                        vtb = trich1.find("B. ")
                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + trich1[:vtb]
                        vtb = vtb + 3
                        trich1 = trich1[vtb:]
                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1[:vtb]
                        #-Xu ly tiep các đáp án sau ở hàng kế tiếp
                        chay += 1
                        if chay < htongdoan and f"(<{h_cau}>)" not in chtml_content[chay] and "[<br>]" not in chtml_content[chay]:
                            s = chtml_content[chay]
                            khop1 = khop_ABCD(s)
                            #if khop1 in [["C."], ["D."]] :
                            if all(item in lssanh for item in khop1):
                            #if (" C. " in s or " c. " in s or r"<u>C.</u>" in s or r"<u>c.</u>" in s or r"<u>C.*</u>" in s or r"<u>c.*</u>" in s):
                                khopans = khop_ans(s)
                                if khopans == ["C."] :
                                #if r"<u>C.</u>" in s or r"<u>c.</u>" in s or r"<u>C.*</u>" in s or r"<u>c.*</u>" in s:
                                    h_dapan = f"{h_cau}C"
                                elif khopans == ["D."] :
                                    h_dapan = f"{h_cau}D"
                                if h_dapan == "":
                                    h_dapan = f"{h_cau}A"
                                trich1 = s.replace("<u>", "")
                                s1 = trich1.replace(r"</u>", "")
                                vta = s1.find("C. ") + 3
                                vtb = s1.find("D. ")
                                sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s1[vta:vtb]
                                vtb = vtb + 1
                                trich1 = s1[vtb:]
                                #-------------------
                                sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1
                                chay += 1
                    elif len(khop) == 1 :
                        # xu ly cau va chi co dap an A cung 1 hang
                        vta = s.find("A.")
                        v0 = s.find(f"{h_cau}>)") + 3#lấy vị trí đầu của câu hỏi con
                        #Nếu vta >=0 thì lấy câu hỏi con trước các đáp án abcd và bao trong cặp [< ]<
                        if vta > 3:
                            sm_hcontents[inchay]["answer"] = "[<" + s[v0:vta] + ">]"
                        #sm_hcontents[inchay]["answer" + f"{h_cau}"] = s[:vta]
                        s = s[vta:]# tra ve s la dap an A
                        # copy phần lấy đáp án vào
                        # Lấy đáp án A
                        while chay < htongdoan:
                            #s = nhtml_content[chay]
                            if "[<br>]" in s or "[<BR>]" in s or f"(<{h_cau + 1}>)" in s:
                                break
                            else:
                                h_dapan = ""
                                khopans = khop_ans(s)
                                if khopans == ["A."] :
                                #if r"<u>A.</u>" in s or r"<u>a.</u>" in s or r"<u>A.*</u>" in s or r"<u>a.*</u>" in s:
                                    h_dapan = f"{h_cau}A"
                                    s = chtml_content[chay]
                                    vt = s.find(r"/u") + 3
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + s[vt:]
                                    chay += 1
                                    while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                        khop1 = khop_ABCD(chtml_content[chay])
                                        #if khop1 == ["B."] :
                                        if all(item in lssanh for item in khop1):
                                        #if ("B. " not in chtml_content[chay] or "b. " not in chtml_content[chay] or r"<u>B.</u>" not in chtml_content[chay] or r"<u>b.</u>" not in chtml_content[chay] or r"<u>B.*</u>" not in chtml_content[chay] or r"<u>b.*</u>" not in chtml_content[chay]):
                                            break
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                        chay += 1
                                    #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                                else:
                                    s = chtml_content[chay]
                                    vt = s.find(". ") + 1
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + s[vt:]
                                    #print("chay = " + f"{chay}")
                                    #print("nhtml_content chay answer 1 = " + hcontents[inchay]["answer" + f"{h_cau}"])
                                    chay += 1
                                    while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                        khop1 = khop_ABCD(chtml_content[chay])
                                        #if khop1 == ["B."] :
                                        if all(item in lssanh for item in khop1):
                                        #if ("B. " not in chtml_content[chay] or "b. " not in chtml_content[chay] or r"<u>B.</u>" not in chtml_content[chay] or r"<u>b.</u>" not in chtml_content[chay] or r"<u>B.*</u>" not in chtml_content[chay] or r"<u>b.*</u>" not in chtml_content[chay]):
                                            break
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                        chay += 1
                                    #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                            break
                        # Lấy đáp án B
                        while chay < htongdoan:
                            s = chtml_content[chay]
                            if "[<br>]" in s or "[<BR>]" in s or f"(<{h_cau + 1}>)" in s:
                                break
                            else:
                                khopans = khop_ans(s)
                                if khopans == ["B."] :
                                #if r"<u>B.</u>" in s or r"<u>b.</u>" in s or r"<u>B.*</u>" in s or r"<u>b.*</u>" in s:
                                    if h_dapan == "":
                                        h_dapan = f"{h_cau}B"
                                    else:
                                        #print('Da co dap an A - trung dap an')
                                        loi = "Da co dap an A - trung dap an"
                                    s = chtml_content[chay]
                                    vt = s.find(r"/u") + 3
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                    chay += 1
                                    while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                        khop1 = khop_ABCD(chtml_content[chay])
                                        #if khop1 == ["C."] :
                                        if all(item in lssanh for item in khop1):
                                        #if ("C. " not in chtml_content[chay] or "c. " not in chtml_content[chay] or r"<u>C.</u>" not in chtml_content[chay] or r"<u>c.</u>" not in chtml_content[chay] or r"<u>C.*</u>" not in chtml_content[chay] or r"<u>c.*</u>" not in chtml_content[chay]):
                                            break
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                        chay += 1
                                    #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                                else:
                                    s = chtml_content[chay]
                                    vt = s.find(". ") + 1
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                    chay += 1
                                    while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                        khop1 = khop_ABCD(chtml_content[chay])
                                        #if khop1 == ["C."] :
                                        if all(item in lssanh for item in khop1):
                                        #if ("C. " not in chtml_content[chay] or "c. " not in chtml_content[chay] or r"<u>C.</u>" not in chtml_content[chay] or r"<u>c.</u>" not in chtml_content[chay] or r"<u>C.*</u>" not in chtml_content[chay] or r"<u>c.*</u>" not in chtml_content[chay]):
                                            break
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                        chay += 1
                                    #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                            break
                        # Lấy đáp án C    
                        while chay < htongdoan:
                            s = chtml_content[chay]
                            if "[<br>]" in s or "[<BR>]" in s or f"(<{h_cau + 1}>)" in s:
                                break
                            else:
                                khopans = khop_ans(s)
                                if khopans == ["C."] :
                                #if r"<u>C.</u>" in s or r"<u>c.</u>" in s or r"<u>C.*</u>" in s or r"<u>c.*</u>" in s:
                                    if h_dapan == "":
                                        h_dapan = f"{h_cau}C"
                                    else:
                                        #print('Da co dap an A, B - trung dap an')
                                        loi = "Da co dap an A, B - trung dap an"
                                    s = chtml_content[chay]
                                    vt = s.find(r"/u") + 3
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                    chay += 1
                                    while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                        khop1 = khop_ABCD(chtml_content[chay])
                                        #if khop1 == ["D."] :
                                        if all(item in lssanh for item in khop1):
                                        #if ("D. " not in chtml_content[chay] or "d. " not in chtml_content[chay] or r"<u>D.</u>" not in chtml_content[chay] or r"<u>d.</u>" not in chtml_content[chay] or r"<u>D.*</u>" not in chtml_content[chay] or r"<u>d.*</u>" not in chtml_content[chay]):
                                            break
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                        chay += 1
                                    #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                                else:
                                    s = chtml_content[chay]
                                    vt = s.find(". ") + 1
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                    chay += 1
                                    while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                        khop1 = khop_ABCD(chtml_content[chay])
                                        #if khop1 == ["D."] :
                                        if all(item in lssanh for item in khop1):
                                        #if ("D. " not in chtml_content[chay] or "d. " not in chtml_content[chay] or r"<u>D.</u>" not in chtml_content[chay] or r"<u>d.</u>" not in chtml_content[chay] or r"<u>D.*</u>" not in chtml_content[chay] or r"<u>d.*</u>" not in chtml_content[chay]):
                                            break
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                        chay += 1
                                    #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                            break
                        # Lấy đáp án D
                        while chay < htongdoan:
                            s = chtml_content[chay]
                            if "[<br>]" in s or "[<BR>]" in s or f"(<{h_cau + 1}>)" in s:
                                break
                            else:
                                khopans = khop_ans(s)
                                if khopans == ["D."] :
                                #if r"<u>D.</u>" in s or r"<u>d.</u>" in s or r"<u>D.*</u>" in s or r"<u>d.*</u>" in s:
                                    if h_dapan == "":
                                        h_dapan = f"{h_cau}D"
                                    else:
                                        #print('Da co dap an A, B, C - trung dap an')
                                        loi = "Da co dap an A, B, C - trung dap an"
                                    s = chtml_content[chay]
                                    vt = s.find(r"/u") + 3
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                    chay += 1
                                    while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                        khop1 = khop_ABCD(chtml_content[chay])
                                        #if khop1 == ["E."] :
                                        if all(item in lssanh for item in khop1):
                                        #if ("E. " not in chtml_content[chay] or "e. " not in chtml_content[chay] or r"<u>E.</u>" not in chtml_content[chay] or r"<u>e.</u>" not in chtml_content[chay] or r"<u>E.*</u>" not in chtml_content[chay] or r"<u>e.*</u>" not in chtml_content[chay]):
                                            break
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                        chay += 1
                                    #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                                else:
                                    s = chtml_content[chay]
                                    vt = s.find(". ") + 1
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                    chay += 1
                                    while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                        khop1 = khop_ABCD(chtml_content[chay])
                                        #if khop1 == ["E."] :
                                        if all(item in lssanh for item in khop1):
                                        #if ("E. " not in chtml_content[chay] or "e. " not in chtml_content[chay] or r"<u>E.</u>" not in chtml_content[chay] or r"<u>e.</u>" not in chtml_content[chay] or r"<u>E.*</u>" not in chtml_content[chay] or r"<u>e.*</u>" not in chtml_content[chay]):
                                            break
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                        chay += 1
                                    #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                            #print("chay = " + f"{chay}")
                            #print("nhtml_content chay answer 1 = " + hcontents[inchay]["answer" + f"{h_cau}"])
                            break

                            # kết thúc phần lấy đáp án vào
                    # Ket thuc phan Thêm vào để xử lý các đáp án nằm cùng dòng với câu hỏi (câu hỏi trống)
                    else:
                        v0 = s.find(f"{h_cau}>)") + 3#lấy vị trí đầu của câu hỏi con
                        ass_cauhoi = "[<" + s[v0:]
                        #sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "[<" + s[v0:]
                        # s là nguyên 1 dòng câu hỏi con ở trên
                        chay += 1
                        s = chtml_content[chay]# tra ve s la dap an A
                        #--Xử lý các đáp án cùng 1 dòng
                        khop = khop_ABCD(s)
                        #aok = ("\tA. " in s or "A. " in s or "a. " in s or "\t<u>A.</u>" in s or r"<u>A.</u>" in s or r"<u>a.</u>" in s or r"<u>A.*</u>" in s or r"<u>a.*</u>" in s)
                        #bok = ("\tB. " in s or "B. " in s or "b. " in s or "\t<u>B.</u>" in s or r"<u>B.</u>" in s or r"<u>b.</u>" in s or r"<u>B.*</u>" in s or r"<u>b.*</u>" in s)
                        #cok = ("\tC. " in s or "C. " in s or "c. " in s or "\t<u>C.</u>" in s or "<u>C.</u>" in s or r"<u>c.</u>" in s or r"<u>C.*</u>" in s or r"<u>c.*</u>" in s)
                        #dok = ("\tD. " in s or "D. " in s or "d. " in s or "\t<u>D.</u>" in s or "<u>D.</u>" in s or r"<u>d.</u>" in s or r"<u>D.*</u>" in s or r"<u>d.*</u>" in s)
                        if len(khop) >= 2 :
                        #if (aok and bok):
                            if len(khop) >= 3 :
                            #if (bok and cok):
                                if len(khop) >= 4 :
                                #if (cok and dok):
                                    # xu ly cau va cac dap an cung 1 hang
                                    khopans = khop_ans(s)
                                    if khopans != [] :
                                        h_dapan = f"{h_cau}" + khopans[0][:1]
                                    else:
                                        h_dapan = f"{h_cau}A"
                                    """
                                    if r"<u>A.</u>" in s or r"<u>a.</u>" in s or r"<u>A.*</u>" in s or r"<u>a.*</u>" in s:
                                        h_dapan = f"{h_cau}A"
                                    elif r"<u>B.</u>" in s or r"<u>b.</u>" in s or r"<u>B.*</u>" in s or r"<u>b.*</u>" in s:
                                        h_dapan = f"{h_cau}B"
                                    elif r"<u>C.</u>" in s or r"<u>c.</u>" in s or r"<u>C.*</u>" in s or r"<u>c.*</u>" in s:
                                        h_dapan = f"{h_cau}C"
                                    elif r"<u>D.</u>" in s or r"<u>d.</u>" in s or r"<u>D.*</u>" in s or r"<u>d.*</u>" in s:
                                        h_dapan = f"{h_cau}D"
                                    elif r"<u>E.</u>" in s or r"<u>e.</u>" in s or r"<u>E.*</u>" in s or r"<u>e.*</u>" in s:
                                        h_dapan = f"{h_cau}E"
                                    elif r"<u>F.</u>" in s or r"<u>f.</u>" in s or r"<u>F.*</u>" in s or r"<u>f.*</u>" in s:
                                        h_dapan = f"{h_cau}F"
                                    """
                                    #print("dap an 1 la " + h_dapan)
                                    trich1 = s.replace("<u>", "")
                                    s1 = trich1.replace(r"</u>", "")
                                    s1 = s1.replace("\t", " ")
                                    vta = s1.find("A. ")
                                    v0 = s1.find(f"{h_cau}>)") + 3#lấy vị trí đầu của câu hỏi con
                                    #Nếu vta >=0 thì lấy câu hỏi con trước các đáp án abcd và bao trong cặp [< ]<
                                    if vta>=0:
                                        ass_cauhoi = ass_cauhoi + s1[v0:vta] + ">]"
                                        #sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "[<" + s1[v0:vta] + ">]"
                                    #--------------------
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + ass_cauhoi
                                    #Sau khi xong câu hỏi con mới đưa vào sm_hcontents
                                    vta = vta + 3
                                    trich1 = s1[vta:]
                                    vtb = trich1.find("B. ")
                                    #print(trich1)
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + trich1[:vtb]
                                    vtb = vtb + 3
                                    trich1 = trich1[vtb:]
                                    #-------------------
                                    vtb = trich1.find("C. ")
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1[:vtb]
                                    vtb = vtb + 3
                                    trich1 = trich1[vtb:]
                                    #-------------------
                                    vtb = trich1.find("D. ")
                                    if vtb == -1:
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1
                                    else:
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1[:vtb]
                                        vtb = vtb + 3
                                        trich1 = trich1[vtb:]
                                    #-------------------
                                    vtb = trich1.find("E. ")
                                    if vtb == -1:
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1
                                    else:
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1[:vtb]
                                        vtb = vtb + 3
                                        trich1 = trich1[vtb:]
                                    chay += 1
                                else:
                                    # xu ly cau va 2 dap an cung 1 hang
                                    khopans = khop_ans(s)
                                    if khopans == ["A."] :
                                    #print("sao lai chay xuong day")
                                    #if r"<u>A.</u>" in s or r"<u>a.</u>" in s or r"<u>A.*</u>" in s or r"<u>a.*</u>" in s:
                                        h_dapan = f"{h_cau}A"
                                    elif khopans == ["B."] :
                                        h_dapan = f"{h_cau}B"
                            
                                    trich1 = s.replace("<u>", "")
                                    s1 = trich1.replace(r"</u>", "")
                                    vta = s1.find("A. ")
                                    v0 = s1.find(f"{h_cau}>)") + 3#lấy vị trí đầu của câu hỏi con
                                    #Nếu vta >=0 thì lấy câu hỏi con trước các đáp án abcd và bao trong cặp [< ]<
                                    if vta>=0:
                                        sm_hcontents[inchay]["answer"] = "[<" + s1[v0:vta] + ">]"
                                    #----------Khuc trên này chưa sửa, chưa hiểu có sửa không vì trước đo ass_cauhoi đã có [< + ... rồi]
                                    vta = vta + 3
                                    trich1 = s1[vta:]
                                    vtb = trich1.find("B. ")
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + trich1[:vtb]
                                    vtb = vtb + 3
                                    trich1 = trich1[vtb:]
                                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1[:vtb]
                                    #-Xu ly tiep các đáp án sau ở hàng kế tiếp
                                    chay += 1
                                    if chay < htongdoan and f"(<{h_cau}>)" not in chtml_content[chay] and "[<br>]" not in chtml_content[chay]:
                                        s = chtml_content[chay]
                                        khop1 = khop_ABCD(s)
                                        #if khop1 in [["C."], ["D."]] :
                                        if all(item in lssanh for item in khop1):
                                        #if ("C. " in s or "c. " in s or r"<u>C.</u>" in s or r"<u>c.</u>" in s or r"<u>C.*</u>" in s or r"<u>c.*</u>" in s):
                                            khopans = khop_ans(s)
                                            if khopans == ["C."] :
                                            #if r"<u>C.</u>" in s or r"<u>c.</u>" in s or r"<u>C.*</u>" in s or r"<u>c.*</u>" in s:
                                                h_dapan = f"{h_cau}C"
                                            elif khopans == ["D."] :
                                                h_dapan = f"{h_cau}D"
                                            if h_dapan == "":
                                                h_dapan = f"{h_cau}A"
                                            trich1 = s.replace("<u>", "")
                                            s1 = trich1.replace(r"</u>", "")
                                            vta = s1.find("C. ") + 3
                                            vtb = s1.find("D. ")
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s1[vta:vtb]
                                            vtb = vtb + 1
                                            trich1 = s1[vtb:]
                                            #-------------------
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + trich1
                                            chay += 1
                        #--Kết thúc xử lý các đáp án cùng 1 dòng và lùi vào 1 tab
                        else:
                            #Phần này xử lý sau câu hỏi con là mỗi đáp án 1 hàng
                            if "[<br>]" in s or "[<BR>]" in s or f"(<{h_cau + 1}>)" in s:
                                break
                            while chay < htongdoan:
                                khop = khop_ABCD(s)
                                #if khop[0] == "A." :
                                if all(item in lssanh for item in khop):
                                #if ("A. " in s or "a. " in s or r"<u>A.</u>" in s or r"<u>a.</u>" in s or r"<u>A.*</u>" in s or r"<u>a.*</u>" in s):
                                    break
                                #sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + s
                                ass_cauhoi = ass_cauhoi + "<br>" + s
                                chay += 1
                                if chay < htongdoan:
                                    s = chtml_content[chay]
                                    if "[<br>]" in s or "[<BR>]" in s or f"(<{h_cau + 1}>)" in s:
                                        #print('Khong co dap an')
                                        loi = "Khong co dap an"
                                        break
                                else:
                                    break
                            #Đã lấy xong phần câu hỏi con
                            ass_cauhoi = ass_cauhoi + ">]"
                            #sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + ">]"
                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + ass_cauhoi
                            # Lấy đáp án A
                            while chay < htongdoan:
                                s = chtml_content[chay]
                                if "[<br>]" in s or "[<BR>]" in s or f"(<{h_cau + 1}>)" in s:
                                    break
                                else:
                                    h_dapan = "" #hcontents[inchay]["correctans"]
                                    khopans = khop_ans(s)
                                    if khopans == ["A."] :
                                    #if r"<u>A.</u>" in s or r"<u>a.</u>" in s or r"<u>A.*</u>" in s or r"<u>a.*</u>" in s:
                                        h_dapan = f"{h_cau}A"
                                        #s = chtml_content[chay]
                                        vt = s.find(r"/u") + 3
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                            khop1 = khop_ABCD(chtml_content[chay])
                                            #if khop1[0] == "B." :
                                            if all(item in lssanh for item in khop1):
                                            #if ("B. " not in chtml_content[chay] or "b. " not in chtml_content[chay] or r"<u>B.</u>" not in chtml_content[chay] or r"<u>b.</u>" not in chtml_content[chay] or r"<u>B.*</u>" not in chtml_content[chay] or r"<u>b.*</u>" not in chtml_content[chay]):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                            chay += 1
                                        #sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "]"
                                    else:
                                        #s = chtml_content[chay]
                                        vt = s.find(".") + 2
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                            khop1 = khop_ABCD(chtml_content[chay])
                                            #if khop1[0] == "B." :
                                            if all(item in lssanh for item in khop1):
                                            #if ("B. " not in chtml_content[chay] or "b. " not in chtml_content[chay] or r"<u>B.</u>" not in chtml_content[chay] or r"<u>b.</u>" not in chtml_content[chay] or r"<u>B.*</u>" not in chtml_content[chay] or r"<u>b.*</u>" not in chtml_content[chay]):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                            chay += 1
                                        #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                                break
                            # Lấy đáp án B
                            while chay < htongdoan:
                                s = chtml_content[chay]
                                if "[<br>]" in s or "[<BR>]" in s or f"(<{h_cau + 1}>)" in s:
                                    break
                                else:
                                    khopans = khop_ans(s)
                                    if khopans == ["B."] :
                                    #if r"<u>B.</u>" in s or r"<u>b.</u>" in s or r"<u>B.*</u>" in s or r"<u>b.*</u>" in s:
                                        if h_dapan == "":
                                            h_dapan = f"{h_cau}B"
                                        else:
                                            #print('Da co dap an A - trung dap an')
                                            loi = "Da co dap an A - trung dap an"
                                        #s = chtml_content[chay]
                                        vt = s.find(r"/u") + 3
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                            khop1 = khop_ABCD(chtml_content[chay])
                                            #if khop1[0] == "C." :
                                            if all(item in lssanh for item in khop1):
                                            #if ("C. " not in chtml_content[chay] or "c. " not in chtml_content[chay] or r"<u>C.</u>" not in chtml_content[chay] or r"<u>c.</u>" not in chtml_content[chay] or r"<u>C.*</u>" not in chtml_content[chay] or r"<u>c.*</u>" not in chtml_content[chay]):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                            chay += 1
                                        #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                                    else:
                                        #s = chtml_content[chay]
                                        vt = s.find(".") + 2
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                            khop1 = khop_ABCD(chtml_content[chay])
                                            #if khop1[0] == "C." :
                                            if all(item in lssanh for item in khop1):
                                            #if ("C. " not in chtml_content[chay] or "c. " not in chtml_content[chay] or r"<u>C.</u>" not in chtml_content[chay] or r"<u>c.</u>" not in chtml_content[chay] or r"<u>C.*</u>" not in chtml_content[chay] or r"<u>c.*</u>" not in chtml_content[chay]):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                            chay += 1
                                        #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                                break
                            # Lấy đáp án C    
                            while chay < htongdoan:
                                s = chtml_content[chay]
                                if "[<br>]" in s or "[<BR>]" in s or f"(<{h_cau + 1}>)" in s:
                                    break
                                else:
                                    khopans = khop_ans(s)
                                    if khopans == ["C."] :
                                    #if r"<u>C.</u>" in s or r"<u>c.</u>" in s or r"<u>C.*</u>" in s or r"<u>c.*</u>" in s:
                                        if h_dapan == "":
                                            h_dapan = f"{h_cau}C"
                                        else:
                                            #print('Da co dap an A, B - trung dap an')
                                            loi = "Da co dap an A, B - trung dap an"
                                        #s = chtml_content[chay]
                                        vt = s.find(r"/u") + 3
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                            khop1 = khop_ABCD(chtml_content[chay])
                                            #if khop1[0] == "D." :
                                            if all(item in lssanh for item in khop1):
                                            #if ("D. " not in chtml_content[chay] or "d. " not in chtml_content[chay] or r"<u>D.</u>" not in chtml_content[chay] or r"<u>d.</u>" not in chtml_content[chay] or r"<u>D.*</u>" not in chtml_content[chay] or r"<u>d.*</u>" not in chtml_content[chay]):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                            chay += 1
                                        #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                                    else:
                                        #s = chtml_content[chay]
                                        vt = s.find(".") + 2
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                            khop1 = khop_ABCD(chtml_content[chay])
                                            #if khop1[0] == "D." :
                                            if all(item in lssanh for item in khop1):
                                            #if ("D. " not in chtml_content[chay] or "d. " not in chtml_content[chay] or r"<u>D.</u>" not in chtml_content[chay] or r"<u>d.</u>" not in chtml_content[chay] or r"<u>D.*</u>" not in chtml_content[chay] or r"<u>d.*</u>" not in chtml_content[chay]):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                            chay += 1
                                        #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                                break
                            # Lấy đáp án D
                            while chay < htongdoan:
                                s = chtml_content[chay]
                                if "[<br>]" in s or "[<BR>]" in s or f"(<{h_cau + 1}>)" in s:
                                    break
                                else:
                                    khopans = khop_ans(s)
                                    if khopans == ["D."] :
                                    #if r"<u>D.</u>" in s or r"<u>d.</u>" in s or r"<u>D</u><u>.</u>" in s or r"<u>d</u><u>.</u>" in s or r"<u>D.*</u>" in s or r"<u>d.*</u>" in s:
                                        if h_dapan == "":
                                            h_dapan = f"{h_cau}D"
                                        else:
                                            #print('Da co dap an A, B, C - trung dap an')
                                            loi = "Da co dap an A, B, C - trung dap an"
                                        #s = chtml_content[chay]
                                        s1 = s.replace(r"<u>", "")
                                        s1 = s1.replace(r"</u>", "")
                                        s1 = s1.replace(r"<\t>", "")
                                        vt = s1.find("D.") + 2
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s1[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                            khop1 = khop_ABCD(chtml_content[chay])
                                            #if khop1[0] == "E." :
                                            if all(item in lssanh for item in khop1):
                                            #if ("E. " not in chtml_content[chay] or "e. " not in chtml_content[chay] or r"<u>E.</u>" not in chtml_content[chay] or r"<u>e.</u>" not in chtml_content[chay] or r"<u>E.*</u>" not in chtml_content[chay] or r"<u>e.*</u>" not in chtml_content[chay]):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                            chay += 1
                                        #sm_hcontents[inchay]["answer" + f"{h_cau}"] = sm_hcontents[inchay]["answer" + f"{h_cau}"] + "]"
                                    else:
                                        #s = chtml_content[chay]
                                        vt = s.find(".") + 2
                                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<##>" + s[vt:]
                                        chay += 1
                                        while (chay < htongdoan and "[<br>]" not in chtml_content[chay] and f"(<{h_cau + 1}>)" not in chtml_content[chay]):
                                            khop1 = khop_ABCD(chtml_content[chay])
                                            #if khop1[0] == "E." :
                                            if all(item in lssanh for item in khop1):
                                            #if ("E. " not in chtml_content[chay] or "e. " not in chtml_content[chay] or r"<u>E.</u>" not in chtml_content[chay] or r"<u>e.</u>" not in chtml_content[chay] or r"<u>E.*</u>" not in chtml_content[chay] or r"<u>e.*</u>" not in chtml_content[chay]):
                                                break
                                            sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<br>" + chtml_content[chay]
                                            chay += 1
                                        #sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "]"
                                #print("chay = " + f"{chay}")
                                #print("nhtml_content chay answer 1 = " + hcontents[inchay]["answer" + f"{h_cau}"])
                                break
                    #-----Dừng-----------------
                    sm_hcontents[inchay]["correctans"] = sm_hcontents[inchay]["correctans"] + h_dapan
                    # Kiem tra truoc khi quay lai vong lap tung cau hoi
                    if chay < htongdoan and f"(<{h_cau + 1}>)" in chtml_content[chay]:
                        h_cau += 1
                        # Thêm phần tách câu trong nhóm, định dạng bằng chữ <#GP#>
                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<#GP#>"
                    elif chay < htongdoan and ("[<br>]" in chtml_content[chay] or "[<BR>]" in chtml_content[chay]):
                        break
                # Kiem tra truoc khi quay lai doan 2
                if (chay < htongdoan and "[<br>]" in chtml_content[chay]) or (chay < htongdoan and "[<BR>]" in chtml_content[chay]):
                    chay += 1
                    tt = tt + 1
                    mch = mmon + f"{tt:04d}"
                    datenum = date.today()
                    datesua = datenum.strftime('%Y-%m-%d')
                    sm_hcontents.append({"macauhoi": mch, "titles": mtd, "questions": "Noi dung cau hoi", "picques": "file hinh cau hoi", "audques": "file amthanh", "clipques": "file video", "permuteques": False,
                        "answer": "Dap an 1<##>Dap an 2<##>Dap an 3", "picans": "", "audans": "", "clipans": "", "permuteans": "T",
                        "classifyques": "LC, DT, DS, NH, GD, TL", "partques": tendvkt, "levelques": mdo, "correctans": "", "layoutans1": 1, "mamon": mmon, "madvkt": mdvkt, "createdate": datesua})
                    
                    inchay += 1
                    sm_hcontents[inchay]["classifyques"] = "NH"
                    #sm_hcontents[inchay]["titles"] = sm_hcontents[inchayfirst]["titles"]
                    sm_hcontents[inchay]["titles"] = mtd

    else:
        loi = "Tài liệu không có đoạn văn bản nào."

    return sm_hcontents, loi

def up_contents_gd(chtml_content, mmon, sott, mdvkt, tendvkt, mtd, mdo):
    loi = ""
    #lisst so sanh
    lssanh = [["A."], ["B."], ["C."], ["D."], ["E."], ["F."], ["G."], ["H."]]
    tt = sott
    mch = mmon + f"{tt:04d}"
    datenum = date.today()
    datesua = datenum.strftime('%Y-%m-%d')
    sm_hcontents = [{"macauhoi": mch, "titles": mtd, "questions": "Noi dung cau hoi", "picques": "file hinh cau hoi", "audques": "file amthanh", "clipques": "file video", "permuteques": True,
            "answer": "Dap an 1<##>Dap an 2<##>Dap an 3", "picans": "", "audans": "", "clipans": "", "permuteans": "T",
            "classifyques": "LC, DT, DS, NH, NL, TL", "partques": tendvkt, "levelques": mdo, "correctans": "", "layoutans1": 1, "mamon": mmon, "madvkt": mdvkt, "createdate": datesua}]
    htongdoan = len(chtml_content)
    if htongdoan > 1:
        inchayfirst = 0
        inchay = inchayfirst
        chay = 0
        if "[<M>]" in chtml_content[chay] or "[<m>]" in chtml_content[chay]:
            #phan viet cho cau hoi ghép đôi
            sm_hcontents[inchay]["classifyques"] = "GD"
            if "[<M>]" in chtml_content[chay]:
                vt = chtml_content[chay].find("[<M>]") + 5
            else:
                vt = chtml_content[chay].find("[<m>]") + 5
            #sm_hcontents[inchay]["titles"] = chtml_content[chay][vt:]
            sm_hcontents[inchay]["titles"] = mtd
            chay += 1
            # Lay noi dung tung doan nhom
            while chay < htongdoan - 2:
                sm_hcontents[inchay]["questions"] = chtml_content[chay]
                chay += 1
                s = chtml_content[chay]
                #print("chay = " + f"{chay}")
                #print("nhtml_content chay answer 1 = " + nhtml_content[chay])
                #thu = "(<1>)" in s
                #print(thu)
                if (chay >= htongdoan - 1) and ("[<br>]" in s or "[<BR>]" in s):
                    break
                while chay < htongdoan - 1 and "(<1>)" not in s:
                    sm_hcontents[inchay]["questions"] = sm_hcontents[inchay]["questions"] + "<br>" + s
                    #print("chay = " + f"{chay}")
                    #print("nhtml_content chay answer 1 = " + nhtml_content[chay])
                    chay += 1
                    if chay < htongdoan:
                        s = chtml_content[chay]
                        #print("chay = " + f"{chay}")
                        #print("nhtml_content chay answer 1 = " + nhtml_content[chay])
                        if "[<br>]" in s or "[<BR>]" in s:
                            loi = "Khong co phần câu hỏi ghép đôi"
                            break
                    else:
                        break
                # Phần tìm câu hỏi và đáp án cho ghép đôi
                h_cau = 1
                # khởi tạo answer = [] sau đó ghép tất cả vào
                sm_hcontents[inchay]["answer"] = ""
                #print("chay = " + f"{chay}")
                sm_hcontents[inchay]["permuteans"] = "T"
                #print("nhtml_content chay answer 1 = " + nhtml_content[chay])
                s = s.replace("\t", " ")
                sthua = ""#lưu chuỗi các đáp án thừa
                while f"(<{h_cau}>)" in s:
                    # Lấy nội dung câu hỏi 1
                    #sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + f"(<{h_cau}>)"
                    vta = s.find(f"(<{h_cau}>)") + 5
                    trich1 = s[vta:]
                    v1 = trich1.find("</td>")# chỉ lấy đến đây thôi, bỏ các ký hiệu bảng
                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + trich1[:v1]
                    vtb = trich1.find("<A>")
                    vtb = vtb + 3
                    trich2 = trich1[vtb:]
                    vta = trich2.find(f"(<{h_cau + 1}>)")
                    v1 = trich2.find("</td>")# chỉ lấy đến đây thôi, bỏ các ký hiệu bảng
                    if vta == -1:
                        #sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "[" + trich1[vtb:] + "]"
                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "[" + trich2[:v1] + "]"
                        sthua = trich2[v1:]
                    else:
                        vta = vta - 1
                        #sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "[" + trich1[vtb:vta] + "]"
                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "[" + trich2[:v1] + "]"
                    #print("cau hoi va dap án = " + hcontents[inchay]["answer" + f"{h_cau}"])
                    # Kiem tra truoc khi quay lai vong lap tung cau hoi
                    if f"(<{h_cau + 1}>)" in s:
                        h_cau += 1
                        #THêm nối các câu bằng ký hiệu <#MP#>
                        sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<#MP#>"
                    else:
                        break
                    #print(hcontents[inchay]["answer" + f"{h_cau}"])
                # Bo sung thêm các đáp án thừa
                while "<A>" in sthua or "<a>" in sthua:
                    # chèn ký hiệu ngắt câu và 0 cho mỗi đáp án thừa
                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "<#MP#>0"
                    vta = sthua.find("<A>") + 3
                    trich = sthua[vta:]
                    vtb = trich.find("<A>")
                    v1 = trich.find("</td>")# chỉ lấy đến đây thôi, bỏ các ký hiệu bảng
                    sm_hcontents[inchay]["answer"] = sm_hcontents[inchay]["answer"] + "[" + trich[:v1] + "]"
                    sthua = trich[v1+1:]
                    
                # Tiếp tục chạy
                chay += 1
                # Kiem tra truoc khi quay lai doan 2
                if (chay < htongdoan and "[<br>]" in chtml_content[chay]) or (chay < htongdoan and "[<BR>]" in chtml_content[chay]):
                    chay += 1
                    tt = tt + 1
                    mch = mmon + f"{tt:04d}"
                    datenum = date.today()
                    datesua = datenum.strftime('%Y-%m-%d')
                    sm_hcontents.append({"macauhoi": mch, "titles": mtd, "questions": "Noi dung cau hoi", "picques": "file hinh cau hoi", "audques": "file amthanh", "clipques": "file video", "permuteques": True,
                        "answer": "Dap an 1<##>Dap an 2<##>Dap an 3", "picans": "", "audans": "", "clipans": "", "permuteans": "T",
                        "classifyques": "LC, DT, DS, NH, GD, TL", "partques": tendvkt, "levelques": mdo, "correctans": "", "layoutans1": 1, "mamon": mmon, "madvkt": mdvkt, "createdate": datesua})
                    inchay += 1
                    sm_hcontents[inchay]["classifyques"] = "GD"
                    #sm_hcontents[inchay]["titles"] = sm_hcontents[inchayfirst]["titles"]
                    sm_hcontents[inchay]["titles"] = mtd
                    #print("chay = " + f"{chay}")
                    #print("nhtml_content chay doan 2 = " + nhtml_content[chay])

    else:
        loi = "Tài liệu không có đoạn văn bản nào."

    return sm_hcontents, loi

def imp_file(file_path, fmamon, thutucauhoi, combo_dvkt, combo_madvkt, combo_matd, umdocb):
    #file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if file_path:
        ummon = fmamon
        usott = thutucauhoi
        #usott = 116
        # lấy giá trị chọn từ combo đvkt
        utendvkt = combo_dvkt
        umdvkt = combo_madvkt
        
        # lấy giá trị chọn từ combo tiêu đề
        umtd = combo_matd
        uloi = ""
        #print (utendvkt)
        #print (umdo)
        hcontents = [{"macauhoi": "000000000010001", "titles": umtd, "questions": "Noi dung cau hoi", "picques": "file hinh cau hoi", "audques": "file amthanh", "clipques": "file video", "permuteques": True,
            "answer": "Dap an 1<##>Dap an 2<##>Dap an 3", "picans": "file hinh dap an 1,<##>$,<##>file 3", "audans": "file audio dap an 1,<##>$,<##>$", "clipans": "file video dap an 1,<##>$,<##>$", "permuteans": "T,T,T",
            "classifyques": "LC, DT, DS, NH, NL, TL", "partques": "Chương câu hỏi", "levelques": "Mức độ hiểu biết", "correctans": "", "layoutans1": 1, "mamon": "mã môn", "madvkt": "mã đvkt", "createdate": "2026-10-01"}]
        nhtml_content = extract_docx_with_formatting(file_path)
        thtongdoan = len(nhtml_content)
        if thtongdoan > 1:
            #tinchayfirst = 0
            #tinchay = tinchayfirst
            tchay = 0
            if "[<S>]" in nhtml_content[tchay] or "[<s>]" in nhtml_content[tchay]:
                hcontents, uloi = up_contents_lc(nhtml_content, ummon, usott, umdvkt, utendvkt, umtd, umdocb)
            elif "[<G>]" in nhtml_content[tchay] or "[<g>]" in nhtml_content[tchay]:
                hcontents, uloi = up_contents_nh(nhtml_content, ummon, usott, umdvkt, utendvkt, umtd, umdocb)
            elif "[<M>]" in nhtml_content[tchay] or "[<m>]" in nhtml_content[tchay]:
                hcontents, uloi = up_contents_gd(nhtml_content, ummon, usott, umdvkt, utendvkt, umtd, umdocb)
            # Lưu vào csdl
            # Chuyển đổi dict thành instance dcontents của Subjects001 trước khi truyền vào bulk_create
            dcontents = [models.Subjects001(**item) for item in hcontents]
            models.Subjects001.objects.bulk_create(dcontents)
            #print(hcontents)
            #Person.objects.bulk_create(objs, batch_size=1000)
          
#kết thúc hàm def imp_file()

def login1(request):
    # Nhập, sửa tên đăng nhập sinh viên
    #danhsachuser = models.Usercs.objects.all().values()
    if 'husername' in request.session:
        del request.session['husername']
    #----------------
    context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'btnSubmit':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            xuser = request.POST.get('hname')
            xuser = xuser.strip()
            xps = request.POST.get('hps')
            if xuser != "":
                # Kiểm tra có trong Usercs và vai trò là gì để chuyển hướng url
                # nếu là admin thì chuyển đến trang admin
                if models.Usercs.objects.filter(tendn=xuser).exists():
                    # Lấy tên, password, vai trò
                    danhsachuser = models.Usercs.objects.filter(tendn=xuser)
                    if xps == danhsachuser[0].matkhau:
                        request.session['husername'] = xuser
                        if danhsachuser[0].role == 0:
                            context0 = {
                                'title': 'Quản lý người dùng, hệ thống',
                                'year' : datetime.now().year,
                                'ten': xuser,
                            }
                            return render(request, 'app/admin_dashboard.html', context0)
                        elif danhsachuser[0].role == 1:
                            context0 = {
                                'title': 'Thiết lập đề thi, ngày thi',
                                'year' : datetime.now().year,
                                'ten': xuser,
                            }
                            return render(request, 'app/teacher_dashboard.html', context0)
                        else:
                            context0 = {
                                'title': 'Bài thi trắc nghiệm',
                                'year' : datetime.now().year,
                                'ten': xuser,
                            }
                            return render(request, 'app/student_dashboard.html', context0)
                    #mật khẩu không đúng
                    else:
                        print("Mật khẩu không đúng")
                else:
                    print("Không đúng tên đăng nhập")
                # chuyển hướng đến một URL mới:
                return HttpResponseRedirect('/login1/')

    else:
        context = {
            'title': 'Trang đăng nhập',
            'year' : datetime.now().year,
        }
        
    return render(request, 'app/login1.html', context)

def home(request):
    """Renders the home page."""
    assert isinstance(request, HttpRequest)
    return render(
        request,
        'app/index.html',
        {
            'title':'Home Page',
            'year':datetime.now().year,
        }
    )

def contact(request):
    """Renders the contact page."""
    assert isinstance(request, HttpRequest)
    return render(
        request,
        'app/contact.html',
        {
            'title':'Contact',
            'message':'Your contact page.',
            'year':datetime.now().year,
        }
    )

def about(request):
    """Renders the about page."""
    assert isinstance(request, HttpRequest)
    return render(
        request,
        'app/about.html',
        {
            'title':'About',
            'message':'Your application description page.',
            'year':datetime.now().year,
        }
    )

"""
def upload_file(request):
    if request.method == 'POST' and request.FILES.get('tenfile'):
        myfile = request.FILES['tenfile']
        
        # 1. Lấy tên file gốc từ trình duyệt
        file_name = myfile.name 
        
        # 2. Lưu file vào thư mục media của dự án
        fs = FileSystemStorage()
        filename = fs.save(myfile.name, myfile)
        
        # 3. Lấy đường dẫn đầy đủ (URL) hoặc đường dẫn tuyệt đối trên server
        file_url = fs.url(filename) # Đường dẫn URL (vận hành trên web)
        # file_path = fs.path(filename) # Đường dẫn tuyệt đối trên ổ cứng server
        
        return render(request, 'upload.html', {
            'file_url': file_url,
            'file_name': file_name
        })
    return render(request, 'upload.html')
"""

def unitsstart(request):
    # Tạo năm học, kỳ thi
    assert isinstance(request, HttpRequest)
    vitri = models.Namhoc.objects.count()
    ds = models.Namhoc.objects.all().values()
    dsnh = []
    for i in range(vitri):
        newdist = {'manh' : ds[i]['manh'], 'namhk1' : ds[i]['namhk1'], 'namhk2' : ds[i]['namhk2']}
        dsnh.append(newdist)
    #----------------
    context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save1':
            # tạo một instance của biểu mẫu và đưa tên môn học trong kỳ:
            context['u_form'] = forms.NamhocForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['u_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                umanh = context['u_form'].cleaned_data['hmanh']
                umahk1 = context['u_form'].cleaned_data['hmahk1']
                umahk2 = context['u_form'].cleaned_data['hmahk2']
                co = False
                # kiểm tra mã có trong danh sách hay chưa, đề phòng trường hợp người dùng nnhập sai mã
                for row in range(len(dsnh)):
                    if umanh == dsnh[row]['manh']:
                        co = True
                        break
                if not co:
                    if umanh == umahk1 + "-" + umahk2:
                        models.Namhoc.objects.create(manh = umanh, namhk1 = umahk1, namhk2 = umahk2)
                
                return HttpResponseRedirect('/unitsstart/')
        elif pressed_button == 'repair1':
            context['u_form'] = forms.NamhocForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['u_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                umanh = context['u_form'].cleaned_data['hmanh']
                umahk1 = context['u_form'].cleaned_data['hmahk1']
                umahk2 = context['u_form'].cleaned_data['hmahk2']
                co = False
                # kiểm tra mã có trong danh sách hay chưa, đề phòng trường hợp người dùng nnhập sai mã
                for row in range(len(dsnh)):
                    if umanh == dsnh[row]['manh']:
                        co = True
                        break
                if co:
                    models.Namhoc.objects.filter(manh = umanh).update(namhk1 = umahk1, namhk2 = umahk2)
                return HttpResponseRedirect('/unitsstart/')
        elif pressed_button == 'delete1':
            context['u_form'] = forms.NamhocForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['u_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                umanh = context['u_form'].cleaned_data['hmanh']
                umahk1 = context['u_form'].cleaned_data['hmahk1']
                umahk2 = context['u_form'].cleaned_data['hmahk2']
                co = False
                # kiểm tra mã có trong danh sách hay chưa, đề phòng trường hợp người dùng nnhập sai mã
                for row in range(len(dsnh)):
                    if umanh == dsnh[row]['manh']:
                        co = True
                        break
                if co:
                    models.Namhoc.objects.filter(manh = umanh).delete()
                return HttpResponseRedirect('/unitsstart/')
        if pressed_button == 'save2':
            context['u_form2'] = forms.NamhochhForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['u_form2'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                umanh = context['u_form2'].cleaned_data['hmanh']
                co = False
                # kiểm tra mã có trong danh sách hay chưa, đề phòng trường hợp người dùng nnhập sai mã
                for row in range(len(dsnh)):
                    if umanh == dsnh[row]['manh']:
                        co = True
                        break
                if co:
                    models.CurrentSetting.objects.filter(tt = 1).update(fkmanh = umanh)
                    #request.session['hnamhochh'] = umanh
                return HttpResponseRedirect('/unitsstart/')

    else:
        context = {
            'title': "Năm học",
            'ten': "Danh mục năm học:",
            'year' : datetime.now().year,
            'data': dsnh,
            'u_form': forms.NamhocForm(),
            'u_form2': forms.NamhochhForm(),
        }
        
    return render(request, 'app/unitsstart.html', context)

# Hàm tạo hoán vị
def khoi_ao_ve_rong():
    s1 = ""
    v1 = ""
    s2 = ""
    v2 = ""
    s3 = ""
    v3 = ""
    s4 = ""
    v4 = ""
    s5 = ""
    v5 = ""
    s6 = ""
    v6 = ""
    s7 = ""
    v7 = ""
    s8 = ""
    v8 = ""
    s9 = ""
    v9 = ""
    return s1, v1, s2, v2, s3, v3, s4, v4, s5, v5, s6, v6, s7, v7, s8, v8, s9, v9

def xao_tron_de_thi(solg, dap_an_dung):
    #print("nhập vào đáp án đung=")
    #print(dap_an_dung)
    nhan_dap_an = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
    #chuyển dap_an_dung từ chữ ABC sang index
    for i, q in enumerate(nhan_dap_an, 0):
        if dap_an_dung == q:
            dap_an_dung_idx = i
            break
        else:
            dap_an_dung_idx = 0 # '0' đang ở vị trí 0 (A)
    # Tạo list các phương án với số lượng đáp án = solg, dạng [0, 1, 2, 3]
    ldapan = []
    for n in range(solg):
        ldapan.append(n)
    danh_sach_dap_an = {"dap_an": ldapan, "dap_an_dung_index": dap_an_dung_idx}
    # Sao chép danh sách đáp án gốc để không làm thay đổi dữ liệu gốc
    #cac_lua_chon = list(item["dap_an"])
    cac_lua_chon = danh_sach_dap_an['dap_an']
    noi_dung_dung = cac_lua_chon[dap_an_dung_idx]
    # Xáo trộn vị trí các đáp án shuffle xáo trộn vị trí các phần tử trong danh sách x ngay tại chỗ
    random.shuffle(cac_lua_chon)
    # Tìm lại vị trí mới (ký hiệu mới A, B, C, D) của đáp án đúng
    vi_tri_moi = cac_lua_chon.index(noi_dung_dung)
    
    dap_an_dung_moi_abcd = nhan_dap_an[vi_tri_moi]
    
    # lấy phương án xáo trộn các đáp án
    #hàm enumerate dùng để duyệt qua một tập hợp (như danh sách, chuỗi) và trả về
    # cả chỉ số (index) lẫn giá trị của từng phần tử. enumerate(iterable, start=0)
    #iterable: Đối tượng cần lặp (list, tuple, chuỗi...); start: Số bắt đầu của chỉ số, mặc định là 0
    pamoi = ""
    #for idx, da in enumerate(q["dap_an"]):
    for idx, da in enumerate(cac_lua_chon):
        #print(f"  {['A', 'B', 'C', 'D'][idx]}. {da}")
        pamoi = pamoi + str(da)
    #print(f"-> Đáp án đúng mới: {q['dap_an_dung']}\n")
    return pamoi, dap_an_dung_moi_abcd

# Hàm lấy dữ liệu, radom, tạo hoán vị và trả về mảng gồm: chuẩn và 10 hoán vị
def taodechuanvhoanvi(smakt, smamon, slistnoidung, smachuan, smahoanvi):
    #slistnoidung là biến chứa dữ liệu ma trận i theo số lượng matran chọn đề, j theo bên trong mỗi matran
    if slistnoidung != []:
        #lấy all câu hỏi của môn học có cùng đvkt và mức độ hb (nếu số câu lớn hơn 0)
        skq_dechuan = [{"madechuan": smachuan, "fmamon": smamon, "fmakt": smakt, "noidungchuan": "", "dapanchuan":"",
                        "mahoanvi1": "102", "noidunghv1": "", "dapanhv1": "", "pahv1": "",
                        "mahoanvi2": "202", "noidunghv2": "", "dapanhv2": "", "pahv2": "",
                        "mahoanvi3": "302", "noidunghv3": "", "dapanhv3": "", "pahv3": "",
                        "mahoanvi4": "402", "noidunghv4": "", "dapanhv4": "", "pahv4": "",
                        "mahoanvi5": "502", "noidunghv5": "", "dapanhv5": "", "pahv5": "",
                        "mahoanvi6": "602", "noidunghv6": "", "dapanhv6": "", "pahv6": "",
                        "mahoanvi7": "702", "noidunghv7": "", "dapanhv7": "", "pahv7": "",
                        "mahoanvi8": "102", "noidunghv8": "", "dapanhv8": "", "pahv8": "",
                        "mahoanvi9": "102", "noidunghv9": "", "dapanhv9": "", "pahv9": "",
                        "pahvcau1":"", "pahvcau2":"", "pahvcau3":"", "pahvcau4":"", "pahvcau5":"",
                        "pahvcau6":"", "pahvcau7":"", "pahvcau8":"", "pahvcau9":""}]
        # Gán các mã hoán vị
        skq_dechuan[0]['mahoanvi1'] = smahoanvi[0]
        skq_dechuan[0]['mahoanvi2'] = smahoanvi[1]
        skq_dechuan[0]['mahoanvi3'] = smahoanvi[2]
        skq_dechuan[0]['mahoanvi4'] = smahoanvi[3]
        skq_dechuan[0]['mahoanvi5'] = smahoanvi[4]
        skq_dechuan[0]['mahoanvi6'] = smahoanvi[5]
        skq_dechuan[0]['mahoanvi7'] = smahoanvi[6]
        skq_dechuan[0]['mahoanvi8'] = smahoanvi[7]
        skq_dechuan[0]['mahoanvi9'] = smahoanvi[8]
        for i in range(len(slistnoidung)):
            # Mức độ 1 rất dễ
            socaurd = int(slistnoidung[i][1])
            if socaurd > 0 :
                #xử lý mã đvkt nếu = 00 thì chuyển sang ""
                xlmadvkt = slistnoidung[i][0]
                if xlmadvkt == "00":
                    xlmadvkt = ""
                #lấy ra và xáo trộn luôn, xong append vào
                kq_truoc = models.Subjects001.objects.filter(levelques = "RD", mamon = smamon, madvkt = xlmadvkt)
                #lấy đủ số lượng rồi append
                solgkq = len(kq_truoc)
                #random để lấy ch: random.sample(population, k): Chọn ra k phần tử độc lập và duy nhất (không trùng lặp) từ danh sách population
                kq_random = random.sample(range(solgkq), socaurd)
                for j in range(len(kq_random)):
                    skq_dechuan[0]['noidungchuan'] = skq_dechuan[0]['noidungchuan'] + "," + kq_truoc[kq_random[j]].macauhoi
                    skq_dechuan[0]['dapanchuan'] = skq_dechuan[0]['dapanchuan'] + "," + kq_truoc[kq_random[j]].correctans
                    #noidungchuan = kq_truoc[j].macauhoi #mã câu hỏi
                    #Chưa xét đến có xáo trộn hay không, đây chỉ mới xuất theo tuần tự
                    #Tạo 10 hoán vị loại LC
                    if kq_truoc[kq_random[j]].classifyques == "LC":
                        # Đếm số các phương án trả lời
                        ss = kq_truoc[kq_random[j]].answer
                        vtriss = ss.find("<##>")
                        while vtriss == 0:
                            ss = ss[4:]
                            vtriss = ss.find("<##>")
                        solgpa = ss.count("<##>") + 1
                        dap_an_pa = kq_truoc[kq_random[j]].correctans
                        if solgpa > 1:
                            #cho ma hv 1
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi1'] = smahoanvi[0]
                            skq_dechuan[0]['noidunghv1'] = skq_dechuan[0]['noidunghv1'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv1'] = skq_dechuan[0]['dapanhv1'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv1'] = skq_dechuan[0]['pahv1'] + "," + pa_xao_tron
                            #cho ma hv 2
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi2'] = smahoanvi[1]
                            skq_dechuan[0]['noidunghv2'] = skq_dechuan[0]['noidunghv2'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv2'] = skq_dechuan[0]['dapanhv2'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv2'] = skq_dechuan[0]['pahv2'] + "," + pa_xao_tron
                            #cho ma hv 3
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi3'] = smahoanvi[2]
                            skq_dechuan[0]['noidunghv3'] = skq_dechuan[0]['noidunghv3'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv3'] = skq_dechuan[0]['dapanhv3'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv3'] = skq_dechuan[0]['pahv3'] + "," + pa_xao_tron
                            #cho ma hv 4
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi4'] = smahoanvi[3]
                            skq_dechuan[0]['noidunghv4'] = skq_dechuan[0]['noidunghv4'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv4'] = skq_dechuan[0]['dapanhv4'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv4'] = skq_dechuan[0]['pahv4'] + "," + pa_xao_tron
                            #cho ma hv 5
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi5'] = smahoanvi[4]
                            skq_dechuan[0]['noidunghv5'] = skq_dechuan[0]['noidunghv5'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv5'] = skq_dechuan[0]['dapanhv5'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv5'] = skq_dechuan[0]['pahv5'] + "," + pa_xao_tron
                            #cho ma hv 6
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi6'] = smahoanvi[5]
                            skq_dechuan[0]['noidunghv6'] = skq_dechuan[0]['noidunghv6'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv6'] = skq_dechuan[0]['dapanhv6'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv6'] = skq_dechuan[0]['pahv6'] + "," + pa_xao_tron
                            #cho ma hv 7
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi7'] = smahoanvi[6]
                            skq_dechuan[0]['noidunghv7'] = skq_dechuan[0]['noidunghv7'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv7'] = skq_dechuan[0]['dapanhv7'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv7'] = skq_dechuan[0]['pahv7'] + "," + pa_xao_tron
                            #cho ma hv 8
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi8'] = smahoanvi[7]
                            skq_dechuan[0]['noidunghv8'] = skq_dechuan[0]['noidunghv8'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv8'] = skq_dechuan[0]['dapanhv8'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv8'] = skq_dechuan[0]['pahv8'] + "," + pa_xao_tron
                            #cho ma hv 9
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi9'] = smahoanvi[8]
                            skq_dechuan[0]['noidunghv9'] = skq_dechuan[0]['noidunghv9'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv9'] = skq_dechuan[0]['dapanhv9'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv9'] = skq_dechuan[0]['pahv9'] + "," + pa_xao_tron
                    #Tạo 10 hoán vị loại NH
                    elif kq_truoc[kq_random[j]].classifyques == "NH":
                        # Đếm số các câu
                        ssdau = kq_truoc[kq_random[j]].answer
                        vtrissdau = ssdau.find("<#GP#>")
                        dap_an_padau = kq_truoc[kq_random[j]].correctans
                        while vtrissdau == 0:
                            ssdau = ssdau[6:]
                            vtrissdau = ssdau.find("<#GP#>")# Lấy vị trí của hết câu hỏi, đáp án con 1
                        
                        solgcaucon = ssdau.count("<#GP#>") + 1# đếm số câu hỏi, đáp án con
                        jj = 0
                        # max hoasn vij ddax copy
                        skq_dechuan[0]['noidunghv1'] = skq_dechuan[0]['noidunghv1'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv2'] = skq_dechuan[0]['noidunghv2'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv3'] = skq_dechuan[0]['noidunghv3'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv4'] = skq_dechuan[0]['noidunghv4'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv5'] = skq_dechuan[0]['noidunghv5'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv6'] = skq_dechuan[0]['noidunghv6'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv7'] = skq_dechuan[0]['noidunghv7'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv8'] = skq_dechuan[0]['noidunghv8'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv9'] = skq_dechuan[0]['noidunghv9'] + "," + kq_truoc[kq_random[j]].macauhoi
                        # Tạo biến phụ dapanhv_con và pahv_con để lưu giữ đáp án có và phương án là: 1B2D3C và <0321><3120>
                        pv_c1, dv_c1, pv_c2, dv_c2, pv_c3, dv_c3, pv_c4, dv_c4, pv_c5, dv_c5, pv_c6, dv_c6, pv_c7, dv_c7, pv_c8, dv_c8, pv_c9, dv_c9 = khoi_ao_ve_rong()
                        for n1 in range(solgcaucon):
                        #while solgcaucon >= 1 and vtrissdau > 1:
                            #Tách câu con và các phương án con
                            #Tách câu con và các phương án con
                            if vtrissdau > 0:
                                sscau = ssdau[:vtrissdau]
                            else:
                                sscau = ssdau
                            ssdau = ssdau[vtrissdau + 6:]
                            vtrissdau = ssdau.find("<#GP#>")#để ý câu con cuối cùng không có dấu này
                            # Đếm số các phương án trả lời
                            #ss = kq_truoc[j].answer
                            #vtriss = ssphuongan.find("<##>")
                            solgpa = sscau.count("<##>") + 1
                            dap_an_pa = dap_an_padau[jj+1:jj+2]
                            #cho ma hv 1
                            
                            if solgpa > 1:
                                #pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c1 = pv_c1 + "<" + pa_xao_tron + ">"
                                dv_c1 = dv_c1 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c2 = pv_c2 + "<" + pa_xao_tron + ">"
                                dv_c2 = dv_c2 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c3 = pv_c3 + "<" + pa_xao_tron + ">"
                                dv_c3 = dv_c3 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c4 = pv_c4 + "<" + pa_xao_tron + ">"
                                dv_c4 = dv_c4 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c5 = pv_c5 + "<" + pa_xao_tron + ">"
                                dv_c5 = dv_c5 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c6 = pv_c6 + "<" + pa_xao_tron + ">"
                                dv_c6 = dv_c6 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c7 = pv_c7 + "<" + pa_xao_tron + ">"
                                dv_c7 = dv_c7 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c8 = pv_c8 + "<" + pa_xao_tron + ">"
                                dv_c8 = dv_c8 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c9 = pv_c9 + "<" + pa_xao_tron + ">"
                                dv_c9 = dv_c9 + dap_an_padau[jj:jj+1] + dap_an_xao_tron

                            jj = jj + 2
                        #Sau khi nạp vào hoán vị các phần từ, gán vào 9 h vi
                        skq_dechuan[0]['dapanhv1'] = skq_dechuan[0]['dapanhv1'] + "," + dv_c1
                        skq_dechuan[0]['pahv1'] = skq_dechuan[0]['pahv1'] + "," + pv_c1
                        skq_dechuan[0]['dapanhv2'] = skq_dechuan[0]['dapanhv2'] + "," + dv_c2
                        skq_dechuan[0]['pahv2'] = skq_dechuan[0]['pahv2'] + "," + pv_c2
                        skq_dechuan[0]['dapanhv3'] = skq_dechuan[0]['dapanhv3'] + "," + dv_c3
                        skq_dechuan[0]['pahv3'] = skq_dechuan[0]['pahv3'] + "," + pv_c3
                        skq_dechuan[0]['dapanhv4'] = skq_dechuan[0]['dapanhv4'] + "," + dv_c4
                        skq_dechuan[0]['pahv4'] = skq_dechuan[0]['pahv4'] + "," + pv_c4
                        skq_dechuan[0]['dapanhv5'] = skq_dechuan[0]['dapanhv5'] + "," + dv_c5
                        skq_dechuan[0]['pahv5'] = skq_dechuan[0]['pahv5'] + "," + pv_c5
                        skq_dechuan[0]['dapanhv6'] = skq_dechuan[0]['dapanhv6'] + "," + dv_c6
                        skq_dechuan[0]['pahv6'] = skq_dechuan[0]['pahv6'] + "," + pv_c6
                        skq_dechuan[0]['dapanhv7'] = skq_dechuan[0]['dapanhv7'] + "," + dv_c7
                        skq_dechuan[0]['pahv7'] = skq_dechuan[0]['pahv7'] + "," + pv_c7
                        skq_dechuan[0]['dapanhv8'] = skq_dechuan[0]['dapanhv8'] + "," + dv_c8
                        skq_dechuan[0]['pahv8'] = skq_dechuan[0]['pahv8'] + "," + pv_c8
                        skq_dechuan[0]['dapanhv9'] = skq_dechuan[0]['dapanhv9'] + "," + dv_c9
                        skq_dechuan[0]['pahv9'] = skq_dechuan[0]['pahv9'] + "," + pv_c9
                        
            # Mức độ 2 dễ
            socaurd = int(slistnoidung[i][2])
            if socaurd > 0 :
                #xử lý mã đvkt nếu = 00 thì chuyển sang ""
                xlmadvkt = slistnoidung[i][0]
                if xlmadvkt == "00":
                    xlmadvkt = ""
                #lấy ra và xáo trộn luôn, xong append vào
                kq_truoc = models.Subjects001.objects.filter(levelques = "DE", mamon = smamon, madvkt = xlmadvkt)
                #kq_dechuan.append(models.Subjects001.objects.filter(levelques = "DE", mamon = hu_mamon, madvkt = slistnoidung[0][0]))
                #lấy đủ số lượng rồi append
                solgkq = len(kq_truoc)
                #random để lấy ch: random.sample(population, k): Chọn ra k phần tử độc lập và duy nhất (không trùng lặp) từ danh sách population
                kq_random = random.sample(range(solgkq), socaurd)
                for j in range(len(kq_random)):
                    skq_dechuan[0]['noidungchuan'] = skq_dechuan[0]['noidungchuan'] + "," + kq_truoc[kq_random[j]].macauhoi
                    skq_dechuan[0]['dapanchuan'] = skq_dechuan[0]['dapanchuan'] + "," + kq_truoc[kq_random[j]].correctans
                    #noidungchuan = kq_truoc[j].macauhoi #mã câu hỏi
                    #Chưa xét đến có xáo trộn hay không, đây chỉ mới xuất theo tuần tự
                    #Tạo 10 hoán vị
                    #Tạo 10 hoán vị loại LC
                    if kq_truoc[kq_random[j]].classifyques == "LC":
                        # Đếm số các phương án trả lời
                        ss = kq_truoc[kq_random[j]].answer
                        vtriss = ss.find("<##>")
                        while vtriss == 0:
                            ss = ss[4:]
                            vtriss = ss.find("<##>")
                        solgpa = ss.count("<##>") + 1
                        dap_an_pa = kq_truoc[kq_random[j]].correctans
                        if solgpa > 1:
                            #cho ma hv 1
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi1'] = smahoanvi[0]
                            skq_dechuan[0]['noidunghv1'] = skq_dechuan[0]['noidunghv1'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv1'] = skq_dechuan[0]['dapanhv1'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv1'] = skq_dechuan[0]['pahv1'] + "," + pa_xao_tron
                            #cho ma hv 2
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi2'] = smahoanvi[1]
                            skq_dechuan[0]['noidunghv2'] = skq_dechuan[0]['noidunghv2'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv2'] = skq_dechuan[0]['dapanhv2'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv2'] = skq_dechuan[0]['pahv2'] + "," + pa_xao_tron
                            #cho ma hv 3
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi3'] = smahoanvi[2]
                            skq_dechuan[0]['noidunghv3'] = skq_dechuan[0]['noidunghv3'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv3'] = skq_dechuan[0]['dapanhv3'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv3'] = skq_dechuan[0]['pahv3'] + "," + pa_xao_tron
                            #cho ma hv 4
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi4'] = smahoanvi[3]
                            skq_dechuan[0]['noidunghv4'] = skq_dechuan[0]['noidunghv4'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv4'] = skq_dechuan[0]['dapanhv4'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv4'] = skq_dechuan[0]['pahv4'] + "," + pa_xao_tron
                            #cho ma hv 5
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi5'] = smahoanvi[4]
                            skq_dechuan[0]['noidunghv5'] = skq_dechuan[0]['noidunghv5'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv5'] = skq_dechuan[0]['dapanhv5'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv5'] = skq_dechuan[0]['pahv5'] + "," + pa_xao_tron
                            #cho ma hv 6
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi6'] = smahoanvi[5]
                            skq_dechuan[0]['noidunghv6'] = skq_dechuan[0]['noidunghv6'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv6'] = skq_dechuan[0]['dapanhv6'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv6'] = skq_dechuan[0]['pahv6'] + "," + pa_xao_tron
                            #cho ma hv 7
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi7'] = smahoanvi[6]
                            skq_dechuan[0]['noidunghv7'] = skq_dechuan[0]['noidunghv7'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv7'] = skq_dechuan[0]['dapanhv7'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv7'] = skq_dechuan[0]['pahv7'] + "," + pa_xao_tron
                            #cho ma hv 8
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi8'] = smahoanvi[7]
                            skq_dechuan[0]['noidunghv8'] = skq_dechuan[0]['noidunghv8'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv8'] = skq_dechuan[0]['dapanhv8'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv8'] = skq_dechuan[0]['pahv8'] + "," + pa_xao_tron
                            #cho ma hv 9
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi9'] = smahoanvi[8]
                            skq_dechuan[0]['noidunghv9'] = skq_dechuan[0]['noidunghv9'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv9'] = skq_dechuan[0]['dapanhv9'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv9'] = skq_dechuan[0]['pahv9'] + "," + pa_xao_tron
                    #Tạo 10 hoán vị loại NH
                    elif kq_truoc[kq_random[j]].classifyques == "NH":
                        # Đếm số các câu
                        ssdau = kq_truoc[kq_random[j]].answer
                        vtrissdau = ssdau.find("<#GP#>")
                        dap_an_padau = kq_truoc[kq_random[j]].correctans
                        while vtrissdau == 0:
                            ssdau = ssdau[6:]
                            vtrissdau = ssdau.find("<#GP#>")# Lấy vị trí của hết câu hỏi, đáp án con 1
                        
                        solgcaucon = ssdau.count("<#GP#>") + 1# đếm số câu hỏi, đáp án con
                        jj = 0
                        #đã chuyển gán mã hoán vị
                        skq_dechuan[0]['noidunghv1'] = skq_dechuan[0]['noidunghv1'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv2'] = skq_dechuan[0]['noidunghv2'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv3'] = skq_dechuan[0]['noidunghv3'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv4'] = skq_dechuan[0]['noidunghv4'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv5'] = skq_dechuan[0]['noidunghv5'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv6'] = skq_dechuan[0]['noidunghv6'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv7'] = skq_dechuan[0]['noidunghv7'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv8'] = skq_dechuan[0]['noidunghv8'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv9'] = skq_dechuan[0]['noidunghv9'] + "," + kq_truoc[kq_random[j]].macauhoi
                        # Tạo biến phụ dapanhv_con và pahv_con để lưu giữ đáp án có và phương án là: 1B2D3C và <0321><3120>
                        pv_c1, dv_c1, pv_c2, dv_c2, pv_c3, dv_c3, pv_c4, dv_c4, pv_c5, dv_c5, pv_c6, dv_c6, pv_c7, dv_c7, pv_c8, dv_c8, pv_c9, dv_c9 = khoi_ao_ve_rong()
                        for n1 in range(solgcaucon):
                        #while solgcaucon >= 1 and vtrissdau > 1:
                            #Tách câu con và các phương án con
                            if vtrissdau > 0:
                                sscau = ssdau[:vtrissdau]
                            else:
                                sscau = ssdau
                            ssdau = ssdau[vtrissdau + 6:]
                            vtrissdau = ssdau.find("<#GP#>")#để ý câu con cuối cùng không có dấu này
                            # Đếm số các phương án trả lời
                            #ss = kq_truoc[j].answer
                            #vtriss = ssphuongan.find("<##>")
                            solgpa = sscau.count("<##>") + 1
                            dap_an_pa = dap_an_padau[jj+1:jj+2]
                            #print("đap án truyen vào")
                            #print(dap_an_pa)
                            #cho ma hv 1
                            
                            if solgpa > 1:
                                #pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c1 = pv_c1 + "<" + pa_xao_tron + ">"
                                dv_c1 = dv_c1 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c2 = pv_c2 + "<" + pa_xao_tron + ">"
                                dv_c2 = dv_c2 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c3 = pv_c3 + "<" + pa_xao_tron + ">"
                                dv_c3 = dv_c3 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c4 = pv_c4 + "<" + pa_xao_tron + ">"
                                dv_c4 = dv_c4 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c5 = pv_c5 + "<" + pa_xao_tron + ">"
                                dv_c5 = dv_c5 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c6 = pv_c6 + "<" + pa_xao_tron + ">"
                                dv_c6 = dv_c6 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c7 = pv_c7 + "<" + pa_xao_tron + ">"
                                dv_c7 = dv_c7 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c8 = pv_c8 + "<" + pa_xao_tron + ">"
                                dv_c8 = dv_c8 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c9 = pv_c9 + "<" + pa_xao_tron + ">"
                                dv_c9 = dv_c9 + dap_an_padau[jj:jj+1] + dap_an_xao_tron

                            jj = jj + 2
                        #Sau khi nạp vào hoán vị các phần từ, gán vào 9 h vi
                        skq_dechuan[0]['dapanhv1'] = skq_dechuan[0]['dapanhv1'] + "," + dv_c1
                        skq_dechuan[0]['pahv1'] = skq_dechuan[0]['pahv1'] + "," + pv_c1
                        skq_dechuan[0]['dapanhv2'] = skq_dechuan[0]['dapanhv2'] + "," + dv_c2
                        skq_dechuan[0]['pahv2'] = skq_dechuan[0]['pahv2'] + "," + pv_c2
                        skq_dechuan[0]['dapanhv3'] = skq_dechuan[0]['dapanhv3'] + "," + dv_c3
                        skq_dechuan[0]['pahv3'] = skq_dechuan[0]['pahv3'] + "," + pv_c3
                        skq_dechuan[0]['dapanhv4'] = skq_dechuan[0]['dapanhv4'] + "," + dv_c4
                        skq_dechuan[0]['pahv4'] = skq_dechuan[0]['pahv4'] + "," + pv_c4
                        skq_dechuan[0]['dapanhv5'] = skq_dechuan[0]['dapanhv5'] + "," + dv_c5
                        skq_dechuan[0]['pahv5'] = skq_dechuan[0]['pahv5'] + "," + pv_c5
                        skq_dechuan[0]['dapanhv6'] = skq_dechuan[0]['dapanhv6'] + "," + dv_c6
                        skq_dechuan[0]['pahv6'] = skq_dechuan[0]['pahv6'] + "," + pv_c6
                        skq_dechuan[0]['dapanhv7'] = skq_dechuan[0]['dapanhv7'] + "," + dv_c7
                        skq_dechuan[0]['pahv7'] = skq_dechuan[0]['pahv7'] + "," + pv_c7
                        skq_dechuan[0]['dapanhv8'] = skq_dechuan[0]['dapanhv8'] + "," + dv_c8
                        skq_dechuan[0]['pahv8'] = skq_dechuan[0]['pahv8'] + "," + pv_c8
                        skq_dechuan[0]['dapanhv9'] = skq_dechuan[0]['dapanhv9'] + "," + dv_c9
                        skq_dechuan[0]['pahv9'] = skq_dechuan[0]['pahv9'] + "," + pv_c9
            # Mức độ 3 trung bình
            socaurd = int(slistnoidung[i][3])
            if socaurd > 0 :
                #xử lý mã đvkt nếu = 00 thì chuyển sang ""
                xlmadvkt = slistnoidung[i][0]
                if xlmadvkt == "00":
                    xlmadvkt = ""
                #lấy ra và xáo trộn luôn, xong append vào
                kq_truoc = models.Subjects001.objects.filter(levelques = "TB", mamon = smamon, madvkt = xlmadvkt)
                #kq_dechuan.append(models.Subjects001.objects.filter(levelques = "TB", mamon = hu_mamon, madvkt = slistnoidung[0][0]))
                #lấy đủ số lượng rồi append
                solgkq = len(kq_truoc)
                #random để lấy ch: random.sample(population, k): Chọn ra k phần tử độc lập và duy nhất (không trùng lặp) từ danh sách population
                kq_random = random.sample(range(solgkq), socaurd)
                for j in range(len(kq_random)):
                    skq_dechuan[0]['noidungchuan'] = skq_dechuan[0]['noidungchuan'] + "," + kq_truoc[kq_random[j]].macauhoi
                    skq_dechuan[0]['dapanchuan'] = skq_dechuan[0]['dapanchuan'] + "," + kq_truoc[kq_random[j]].correctans
                    #noidungchuan = kq_truoc[j].macauhoi #mã câu hỏi
                    #Chưa xét đến có xáo trộn hay không, đây chỉ mới xuất theo tuần tự
                    #Tạo 10 hoán vị loại LC
                    if kq_truoc[kq_random[j]].classifyques == "LC":
                        # Đếm số các phương án trả lời
                        ss = kq_truoc[kq_random[j]].answer
                        vtriss = ss.find("<##>")
                        while vtriss == 0:
                            ss = ss[4:]
                            vtriss = ss.find("<##>")
                        solgpa = ss.count("<##>") + 1
                        dap_an_pa = kq_truoc[kq_random[j]].correctans
                        if solgpa > 1:
                            #cho ma hv 1
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi1'] = smahoanvi[0]
                            skq_dechuan[0]['noidunghv1'] = skq_dechuan[0]['noidunghv1'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv1'] = skq_dechuan[0]['dapanhv1'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv1'] = skq_dechuan[0]['pahv1'] + "," + pa_xao_tron
                            #cho ma hv 2
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi2'] = smahoanvi[1]
                            skq_dechuan[0]['noidunghv2'] = skq_dechuan[0]['noidunghv2'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv2'] = skq_dechuan[0]['dapanhv2'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv2'] = skq_dechuan[0]['pahv2'] + "," + pa_xao_tron
                            #cho ma hv 3
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi3'] = smahoanvi[2]
                            skq_dechuan[0]['noidunghv3'] = skq_dechuan[0]['noidunghv3'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv3'] = skq_dechuan[0]['dapanhv3'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv3'] = skq_dechuan[0]['pahv3'] + "," + pa_xao_tron
                            #cho ma hv 4
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi4'] = smahoanvi[3]
                            skq_dechuan[0]['noidunghv4'] = skq_dechuan[0]['noidunghv4'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv4'] = skq_dechuan[0]['dapanhv4'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv4'] = skq_dechuan[0]['pahv4'] + "," + pa_xao_tron
                            #cho ma hv 5
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi5'] = smahoanvi[4]
                            skq_dechuan[0]['noidunghv5'] = skq_dechuan[0]['noidunghv5'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv5'] = skq_dechuan[0]['dapanhv5'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv5'] = skq_dechuan[0]['pahv5'] + "," + pa_xao_tron
                            #cho ma hv 6
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi6'] = smahoanvi[5]
                            skq_dechuan[0]['noidunghv6'] = skq_dechuan[0]['noidunghv6'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv6'] = skq_dechuan[0]['dapanhv6'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv6'] = skq_dechuan[0]['pahv6'] + "," + pa_xao_tron
                            #cho ma hv 7
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi7'] = smahoanvi[6]
                            skq_dechuan[0]['noidunghv7'] = skq_dechuan[0]['noidunghv7'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv7'] = skq_dechuan[0]['dapanhv7'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv7'] = skq_dechuan[0]['pahv7'] + "," + pa_xao_tron
                            #cho ma hv 8
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi8'] = smahoanvi[7]
                            skq_dechuan[0]['noidunghv8'] = skq_dechuan[0]['noidunghv8'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv8'] = skq_dechuan[0]['dapanhv8'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv8'] = skq_dechuan[0]['pahv8'] + "," + pa_xao_tron
                            #cho ma hv 9
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi9'] = smahoanvi[8]
                            skq_dechuan[0]['noidunghv9'] = skq_dechuan[0]['noidunghv9'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv9'] = skq_dechuan[0]['dapanhv9'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv9'] = skq_dechuan[0]['pahv9'] + "," + pa_xao_tron
                    #Tạo 10 hoán vị loại NH
                    elif kq_truoc[kq_random[j]].classifyques == "NH":
                        # Đếm số các câu
                        ssdau = kq_truoc[kq_random[j]].answer
                        vtrissdau = ssdau.find("<#GP#>")
                        dap_an_padau = kq_truoc[kq_random[j]].correctans
                        while vtrissdau == 0:
                            ssdau = ssdau[6:]
                            vtrissdau = ssdau.find("<#GP#>")# Lấy vị trí của hết câu hỏi, đáp án con 1
                        
                        solgcaucon = ssdau.count("<#GP#>") + 1# đếm số câu hỏi, đáp án con
                        jj = 0
                        
                        skq_dechuan[0]['noidunghv1'] = skq_dechuan[0]['noidunghv1'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv2'] = skq_dechuan[0]['noidunghv2'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv3'] = skq_dechuan[0]['noidunghv3'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv4'] = skq_dechuan[0]['noidunghv4'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv5'] = skq_dechuan[0]['noidunghv5'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv6'] = skq_dechuan[0]['noidunghv6'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv7'] = skq_dechuan[0]['noidunghv7'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv8'] = skq_dechuan[0]['noidunghv8'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv9'] = skq_dechuan[0]['noidunghv9'] + "," + kq_truoc[kq_random[j]].macauhoi
                        # Tạo biến phụ dapanhv_con và pahv_con để lưu giữ đáp án có và phương án là: 1B2D3C và <0321><3120>
                        pv_c1, dv_c1, pv_c2, dv_c2, pv_c3, dv_c3, pv_c4, dv_c4, pv_c5, dv_c5, pv_c6, dv_c6, pv_c7, dv_c7, pv_c8, dv_c8, pv_c9, dv_c9 = khoi_ao_ve_rong()
                        for n1 in range(solgcaucon):
                        #while solgcaucon >= 1 and vtrissdau > 1:
                            #Tách câu con và các phương án con
                            if vtrissdau > 0:
                                sscau = ssdau[:vtrissdau]
                            else:
                                sscau = ssdau
                            ssdau = ssdau[vtrissdau + 6:]
                            vtrissdau = ssdau.find("<#GP#>")#để ý câu con cuối cùng không có dấu này
                            # Đếm số các phương án trả lời
                            #ss = kq_truoc[j].answer
                            #vtriss = ssphuongan.find("<##>")
                            solgpa = sscau.count("<##>") + 1
                            dap_an_pa = dap_an_padau[jj+1:jj+2]
                            #print("đap án truyen vào")
                            #print(dap_an_pa)
                            #cho ma hv 1
                            
                            if solgpa > 1:
                                #pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c1 = pv_c1 + "<" + pa_xao_tron + ">"
                                dv_c1 = dv_c1 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c2 = pv_c2 + "<" + pa_xao_tron + ">"
                                dv_c2 = dv_c2 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c3 = pv_c3 + "<" + pa_xao_tron + ">"
                                dv_c3 = dv_c3 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c4 = pv_c4 + "<" + pa_xao_tron + ">"
                                dv_c4 = dv_c4 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c5 = pv_c5 + "<" + pa_xao_tron + ">"
                                dv_c5 = dv_c5 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c6 = pv_c6 + "<" + pa_xao_tron + ">"
                                dv_c6 = dv_c6 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c7 = pv_c7 + "<" + pa_xao_tron + ">"
                                dv_c7 = dv_c7 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c8 = pv_c8 + "<" + pa_xao_tron + ">"
                                dv_c8 = dv_c8 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c9 = pv_c9 + "<" + pa_xao_tron + ">"
                                dv_c9 = dv_c9 + dap_an_padau[jj:jj+1] + dap_an_xao_tron

                            jj = jj + 2
                        #Sau khi nạp vào hoán vị các phần từ, gán vào 9 h vi
                        skq_dechuan[0]['dapanhv1'] = skq_dechuan[0]['dapanhv1'] + "," + dv_c1
                        skq_dechuan[0]['pahv1'] = skq_dechuan[0]['pahv1'] + "," + pv_c1
                        skq_dechuan[0]['dapanhv2'] = skq_dechuan[0]['dapanhv2'] + "," + dv_c2
                        skq_dechuan[0]['pahv2'] = skq_dechuan[0]['pahv2'] + "," + pv_c2
                        skq_dechuan[0]['dapanhv3'] = skq_dechuan[0]['dapanhv3'] + "," + dv_c3
                        skq_dechuan[0]['pahv3'] = skq_dechuan[0]['pahv3'] + "," + pv_c3
                        skq_dechuan[0]['dapanhv4'] = skq_dechuan[0]['dapanhv4'] + "," + dv_c4
                        skq_dechuan[0]['pahv4'] = skq_dechuan[0]['pahv4'] + "," + pv_c4
                        skq_dechuan[0]['dapanhv5'] = skq_dechuan[0]['dapanhv5'] + "," + dv_c5
                        skq_dechuan[0]['pahv5'] = skq_dechuan[0]['pahv5'] + "," + pv_c5
                        skq_dechuan[0]['dapanhv6'] = skq_dechuan[0]['dapanhv6'] + "," + dv_c6
                        skq_dechuan[0]['pahv6'] = skq_dechuan[0]['pahv6'] + "," + pv_c6
                        skq_dechuan[0]['dapanhv7'] = skq_dechuan[0]['dapanhv7'] + "," + dv_c7
                        skq_dechuan[0]['pahv7'] = skq_dechuan[0]['pahv7'] + "," + pv_c7
                        skq_dechuan[0]['dapanhv8'] = skq_dechuan[0]['dapanhv8'] + "," + dv_c8
                        skq_dechuan[0]['pahv8'] = skq_dechuan[0]['pahv8'] + "," + pv_c8
                        skq_dechuan[0]['dapanhv9'] = skq_dechuan[0]['dapanhv9'] + "," + dv_c9
                        skq_dechuan[0]['pahv9'] = skq_dechuan[0]['pahv9'] + "," + pv_c9
            # Mức độ 4 khó
            socaurd = int(slistnoidung[i][4])
            if socaurd > 0 :
                #xử lý mã đvkt nếu = 00 thì chuyển sang ""
                xlmadvkt = slistnoidung[i][0]
                if xlmadvkt == "00":
                    xlmadvkt = ""
                #lấy ra và xáo trộn luôn, xong append vào
                kq_truoc = models.Subjects001.objects.filter(levelques = "KH", mamon = smamon, madvkt = xlmadvkt)
                #kq_dechuan.append(models.Subjects001.objects.filter(levelques = "KH", mamon = hu_mamon, madvkt = slistnoidung[0][0]))
                #lấy đủ số lượng rồi append
                solgkq = len(kq_truoc)
                #random để lấy ch: random.sample(population, k): Chọn ra k phần tử độc lập và duy nhất (không trùng lặp) từ danh sách population
                kq_random = random.sample(range(solgkq), socaurd)
                #print(kq_random)
                for j in range(len(kq_random)):
                    skq_dechuan[0]['noidungchuan'] = skq_dechuan[0]['noidungchuan'] + "," + kq_truoc[kq_random[j]].macauhoi
                    skq_dechuan[0]['dapanchuan'] = skq_dechuan[0]['dapanchuan'] + "," + kq_truoc[kq_random[j]].correctans
                    #noidungchuan = kq_truoc[j].macauhoi #mã câu hỏi
                    #Chưa xét đến có xáo trộn hay không, đây chỉ mới xuất theo tuần tự
                    #Tạo 10 hoán vị loại LC
                    if kq_truoc[kq_random[j]].classifyques == "LC":
                        # Đếm số các phương án trả lời
                        ss = kq_truoc[kq_random[j]].answer
                        vtriss = ss.find("<##>")
                        while vtriss == 0:
                            ss = ss[4:]
                            vtriss = ss.find("<##>")
                        solgpa = ss.count("<##>") + 1
                        dap_an_pa = kq_truoc[kq_random[j]].correctans
                        if solgpa > 1:
                            #cho ma hv 1
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi1'] = smahoanvi[0]
                            skq_dechuan[0]['noidunghv1'] = skq_dechuan[0]['noidunghv1'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv1'] = skq_dechuan[0]['dapanhv1'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv1'] = skq_dechuan[0]['pahv1'] + "," + pa_xao_tron
                            #cho ma hv 2
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi2'] = smahoanvi[1]
                            skq_dechuan[0]['noidunghv2'] = skq_dechuan[0]['noidunghv2'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv2'] = skq_dechuan[0]['dapanhv2'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv2'] = skq_dechuan[0]['pahv2'] + "," + pa_xao_tron
                            #cho ma hv 3
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi3'] = smahoanvi[2]
                            skq_dechuan[0]['noidunghv3'] = skq_dechuan[0]['noidunghv3'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv3'] = skq_dechuan[0]['dapanhv3'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv3'] = skq_dechuan[0]['pahv3'] + "," + pa_xao_tron
                            #cho ma hv 4
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi4'] = smahoanvi[3]
                            skq_dechuan[0]['noidunghv4'] = skq_dechuan[0]['noidunghv4'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv4'] = skq_dechuan[0]['dapanhv4'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv4'] = skq_dechuan[0]['pahv4'] + "," + pa_xao_tron
                            #cho ma hv 5
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi5'] = smahoanvi[4]
                            skq_dechuan[0]['noidunghv5'] = skq_dechuan[0]['noidunghv5'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv5'] = skq_dechuan[0]['dapanhv5'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv5'] = skq_dechuan[0]['pahv5'] + "," + pa_xao_tron
                            #cho ma hv 6
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi6'] = smahoanvi[5]
                            skq_dechuan[0]['noidunghv6'] = skq_dechuan[0]['noidunghv6'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv6'] = skq_dechuan[0]['dapanhv6'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv6'] = skq_dechuan[0]['pahv6'] + "," + pa_xao_tron
                            #cho ma hv 7
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi7'] = smahoanvi[6]
                            skq_dechuan[0]['noidunghv7'] = skq_dechuan[0]['noidunghv7'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv7'] = skq_dechuan[0]['dapanhv7'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv7'] = skq_dechuan[0]['pahv7'] + "," + pa_xao_tron
                            #cho ma hv 8
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi8'] = smahoanvi[7]
                            skq_dechuan[0]['noidunghv8'] = skq_dechuan[0]['noidunghv8'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv8'] = skq_dechuan[0]['dapanhv8'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv8'] = skq_dechuan[0]['pahv8'] + "," + pa_xao_tron
                            #cho ma hv 9
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi9'] = smahoanvi[8]
                            skq_dechuan[0]['noidunghv9'] = skq_dechuan[0]['noidunghv9'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv9'] = skq_dechuan[0]['dapanhv9'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv9'] = skq_dechuan[0]['pahv9'] + "," + pa_xao_tron
                    #Tạo 10 hoán vị loại NH
                    elif kq_truoc[kq_random[j]].classifyques == "NH":
                        # Đếm số các câu
                        ssdau = kq_truoc[kq_random[j]].answer
                        vtrissdau = ssdau.find("<#GP#>")
                        dap_an_padau = kq_truoc[kq_random[j]].correctans
                        while vtrissdau == 0:
                            ssdau = ssdau[6:]
                            vtrissdau = ssdau.find("<#GP#>")# Lấy vị trí của hết câu hỏi, đáp án con 1
                        
                        solgcaucon = ssdau.count("<#GP#>") + 1# đếm số câu hỏi, đáp án con
                        jj = 0
                        
                        skq_dechuan[0]['noidunghv1'] = skq_dechuan[0]['noidunghv1'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv2'] = skq_dechuan[0]['noidunghv2'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv3'] = skq_dechuan[0]['noidunghv3'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv4'] = skq_dechuan[0]['noidunghv4'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv5'] = skq_dechuan[0]['noidunghv5'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv6'] = skq_dechuan[0]['noidunghv6'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv7'] = skq_dechuan[0]['noidunghv7'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv8'] = skq_dechuan[0]['noidunghv8'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv9'] = skq_dechuan[0]['noidunghv9'] + "," + kq_truoc[kq_random[j]].macauhoi
                        # Tạo biến phụ dapanhv_con và pahv_con để lưu giữ đáp án có và phương án là: 1B2D3C và <0321><3120>
                        pv_c1, dv_c1, pv_c2, dv_c2, pv_c3, dv_c3, pv_c4, dv_c4, pv_c5, dv_c5, pv_c6, dv_c6, pv_c7, dv_c7, pv_c8, dv_c8, pv_c9, dv_c9 = khoi_ao_ve_rong()
                        for n1 in range(solgcaucon):
                        #while solgcaucon >= 1 and vtrissdau > 1:
                            #Tách câu con và các phương án con
                            if vtrissdau > 0:
                                sscau = ssdau[:vtrissdau]
                            else:
                                sscau = ssdau
                            ssdau = ssdau[vtrissdau + 6:]
                            vtrissdau = ssdau.find("<#GP#>")#để ý câu con cuối cùng không có dấu này
                            # Đếm số các phương án trả lời
                            #ss = kq_truoc[j].answer
                            #vtriss = ssphuongan.find("<##>")
                            solgpa = sscau.count("<##>") + 1
                            dap_an_pa = dap_an_padau[jj+1:jj+2]
                            #print("đap án truyen vào")
                            #print(dap_an_pa)
                            #cho ma hv 1
                            
                            if solgpa > 1:
                                #pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c1 = pv_c1 + "<" + pa_xao_tron + ">"
                                dv_c1 = dv_c1 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c2 = pv_c2 + "<" + pa_xao_tron + ">"
                                dv_c2 = dv_c2 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c3 = pv_c3 + "<" + pa_xao_tron + ">"
                                dv_c3 = dv_c3 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c4 = pv_c4 + "<" + pa_xao_tron + ">"
                                dv_c4 = dv_c4 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c5 = pv_c5 + "<" + pa_xao_tron + ">"
                                dv_c5 = dv_c5 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c6 = pv_c6 + "<" + pa_xao_tron + ">"
                                dv_c6 = dv_c6 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c7 = pv_c7 + "<" + pa_xao_tron + ">"
                                dv_c7 = dv_c7 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c8 = pv_c8 + "<" + pa_xao_tron + ">"
                                dv_c8 = dv_c8 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c9 = pv_c9 + "<" + pa_xao_tron + ">"
                                dv_c9 = dv_c9 + dap_an_padau[jj:jj+1] + dap_an_xao_tron

                            jj = jj + 2
                        #Sau khi nạp vào hoán vị các phần từ, gán vào 9 h vi
                        skq_dechuan[0]['dapanhv1'] = skq_dechuan[0]['dapanhv1'] + "," + dv_c1
                        skq_dechuan[0]['pahv1'] = skq_dechuan[0]['pahv1'] + "," + pv_c1
                        skq_dechuan[0]['dapanhv2'] = skq_dechuan[0]['dapanhv2'] + "," + dv_c2
                        skq_dechuan[0]['pahv2'] = skq_dechuan[0]['pahv2'] + "," + pv_c2
                        skq_dechuan[0]['dapanhv3'] = skq_dechuan[0]['dapanhv3'] + "," + dv_c3
                        skq_dechuan[0]['pahv3'] = skq_dechuan[0]['pahv3'] + "," + pv_c3
                        skq_dechuan[0]['dapanhv4'] = skq_dechuan[0]['dapanhv4'] + "," + dv_c4
                        skq_dechuan[0]['pahv4'] = skq_dechuan[0]['pahv4'] + "," + pv_c4
                        skq_dechuan[0]['dapanhv5'] = skq_dechuan[0]['dapanhv5'] + "," + dv_c5
                        skq_dechuan[0]['pahv5'] = skq_dechuan[0]['pahv5'] + "," + pv_c5
                        skq_dechuan[0]['dapanhv6'] = skq_dechuan[0]['dapanhv6'] + "," + dv_c6
                        skq_dechuan[0]['pahv6'] = skq_dechuan[0]['pahv6'] + "," + pv_c6
                        skq_dechuan[0]['dapanhv7'] = skq_dechuan[0]['dapanhv7'] + "," + dv_c7
                        skq_dechuan[0]['pahv7'] = skq_dechuan[0]['pahv7'] + "," + pv_c7
                        skq_dechuan[0]['dapanhv8'] = skq_dechuan[0]['dapanhv8'] + "," + dv_c8
                        skq_dechuan[0]['pahv8'] = skq_dechuan[0]['pahv8'] + "," + pv_c8
                        skq_dechuan[0]['dapanhv9'] = skq_dechuan[0]['dapanhv9'] + "," + dv_c9
                        skq_dechuan[0]['pahv9'] = skq_dechuan[0]['pahv9'] + "," + pv_c9
            # Mức độ 5 rất khó
            socaurd = int(slistnoidung[i][5])
            if socaurd > 0 :
                #xử lý mã đvkt nếu = 00 thì chuyển sang ""
                xlmadvkt = slistnoidung[i][0]
                if xlmadvkt == "00":
                    xlmadvkt = ""
                #lấy ra và xáo trộn luôn, xong append vào
                kq_truoc = models.Subjects001.objects.filter(levelques = "RK", mamon = smamon, madvkt = xlmadvkt)
                #kq_dechuan.append(models.Subjects001.objects.filter(levelques = "RK", mamon = hu_mamon, madvkt = listnoidung[0][0]))
                #lấy đủ số lượng rồi append
                solgkq = len(kq_truoc)
                #random để lấy ch: random.sample(population, k): Chọn ra k phần tử độc lập và duy nhất (không trùng lặp) từ danh sách population
                kq_random = random.sample(range(solgkq), socaurd)
                for j in range(len(kq_random)):
                    skq_dechuan[0]['noidungchuan'] = skq_dechuan[0]['noidungchuan'] + "," + kq_truoc[kq_random[j]].macauhoi
                    skq_dechuan[0]['dapanchuan'] = skq_dechuan[0]['dapanchuan'] + "," + kq_truoc[kq_random[j]].correctans
                    #noidungchuan = kq_truoc[j].macauhoi #mã câu hỏi
                    #Chưa xét đến có xáo trộn hay không, đây chỉ mới xuất theo tuần tự
                    #Tạo 10 hoán vị loại LC
                    if kq_truoc[kq_random[j]].classifyques == "LC":
                        # Đếm số các phương án trả lời
                        ss = kq_truoc[kq_random[j]].answer
                        vtriss = ss.find("<##>")
                        while vtriss == 0:
                            ss = ss[4:]
                            vtriss = ss.find("<##>")
                        solgpa = ss.count("<##>") + 1
                        dap_an_pa = kq_truoc[kq_random[j]].correctans
                        if solgpa > 1:
                            #cho ma hv 1
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi1'] = smahoanvi[0]
                            skq_dechuan[0]['noidunghv1'] = skq_dechuan[0]['noidunghv1'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv1'] = skq_dechuan[0]['dapanhv1'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv1'] = skq_dechuan[0]['pahv1'] + "," + pa_xao_tron
                            #cho ma hv 2
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi2'] = smahoanvi[1]
                            skq_dechuan[0]['noidunghv2'] = skq_dechuan[0]['noidunghv2'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv2'] = skq_dechuan[0]['dapanhv2'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv2'] = skq_dechuan[0]['pahv2'] + "," + pa_xao_tron
                            #cho ma hv 3
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi3'] = smahoanvi[2]
                            skq_dechuan[0]['noidunghv3'] = skq_dechuan[0]['noidunghv3'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv3'] = skq_dechuan[0]['dapanhv3'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv3'] = skq_dechuan[0]['pahv3'] + "," + pa_xao_tron
                            #cho ma hv 4
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi4'] = smahoanvi[3]
                            skq_dechuan[0]['noidunghv4'] = skq_dechuan[0]['noidunghv4'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv4'] = skq_dechuan[0]['dapanhv4'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv4'] = skq_dechuan[0]['pahv4'] + "," + pa_xao_tron
                            #cho ma hv 5
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi5'] = smahoanvi[4]
                            skq_dechuan[0]['noidunghv5'] = skq_dechuan[0]['noidunghv5'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv5'] = skq_dechuan[0]['dapanhv5'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv5'] = skq_dechuan[0]['pahv5'] + "," + pa_xao_tron
                            #cho ma hv 6
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi6'] = smahoanvi[5]
                            skq_dechuan[0]['noidunghv6'] = skq_dechuan[0]['noidunghv6'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv6'] = skq_dechuan[0]['dapanhv6'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv6'] = skq_dechuan[0]['pahv6'] + "," + pa_xao_tron
                            #cho ma hv 7
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi7'] = smahoanvi[6]
                            skq_dechuan[0]['noidunghv7'] = skq_dechuan[0]['noidunghv7'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv7'] = skq_dechuan[0]['dapanhv7'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv7'] = skq_dechuan[0]['pahv7'] + "," + pa_xao_tron
                            #cho ma hv 8
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi8'] = smahoanvi[7]
                            skq_dechuan[0]['noidunghv8'] = skq_dechuan[0]['noidunghv8'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv8'] = skq_dechuan[0]['dapanhv8'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv8'] = skq_dechuan[0]['pahv8'] + "," + pa_xao_tron
                            #cho ma hv 9
                            pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                            #skq_dechuan[0]['mahoanvi9'] = smahoanvi[8]
                            skq_dechuan[0]['noidunghv9'] = skq_dechuan[0]['noidunghv9'] + "," + kq_truoc[kq_random[j]].macauhoi
                            skq_dechuan[0]['dapanhv9'] = skq_dechuan[0]['dapanhv9'] + "," + dap_an_xao_tron
                            skq_dechuan[0]['pahv9'] = skq_dechuan[0]['pahv9'] + "," + pa_xao_tron
                    #Tạo 10 hoán vị loại NH
                    elif kq_truoc[kq_random[j]].classifyques == "NH":
                        # Đếm số các câu
                        ssdau = kq_truoc[kq_random[j]].answer
                        vtrissdau = ssdau.find("<#GP#>")
                        dap_an_padau = kq_truoc[kq_random[j]].correctans
                        while vtrissdau == 0:
                            ssdau = ssdau[6:]
                            vtrissdau = ssdau.find("<#GP#>")# Lấy vị trí của hết câu hỏi, đáp án con 1
                        
                        solgcaucon = ssdau.count("<#GP#>") + 1# đếm số câu hỏi, đáp án con
                        jj = 0
                        
                        skq_dechuan[0]['noidunghv1'] = skq_dechuan[0]['noidunghv1'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv2'] = skq_dechuan[0]['noidunghv2'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv3'] = skq_dechuan[0]['noidunghv3'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv4'] = skq_dechuan[0]['noidunghv4'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv5'] = skq_dechuan[0]['noidunghv5'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv6'] = skq_dechuan[0]['noidunghv6'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv7'] = skq_dechuan[0]['noidunghv7'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv8'] = skq_dechuan[0]['noidunghv8'] + "," + kq_truoc[kq_random[j]].macauhoi
                        skq_dechuan[0]['noidunghv9'] = skq_dechuan[0]['noidunghv9'] + "," + kq_truoc[kq_random[j]].macauhoi
                        # Tạo biến phụ dapanhv_con và pahv_con để lưu giữ đáp án có và phương án là: 1B2D3C và <0321><3120>
                        pv_c1, dv_c1, pv_c2, dv_c2, pv_c3, dv_c3, pv_c4, dv_c4, pv_c5, dv_c5, pv_c6, dv_c6, pv_c7, dv_c7, pv_c8, dv_c8, pv_c9, dv_c9 = khoi_ao_ve_rong()
                        for n1 in range(solgcaucon):
                        #while solgcaucon >= 1 and vtrissdau > 1:
                            #Tách câu con và các phương án con
                            if vtrissdau > 0:
                                sscau = ssdau[:vtrissdau]
                            else:
                                sscau = ssdau
                            ssdau = ssdau[vtrissdau + 6:]
                            vtrissdau = ssdau.find("<#GP#>")#để ý câu con cuối cùng không có dấu này
                            # Đếm số các phương án trả lời
                            solgpa = sscau.count("<##>") + 1
                            dap_an_pa = dap_an_padau[jj+1:jj+2]
                            #print("đap án truyen vào")
                            #cho ma hv 1
                            if solgpa > 1:
                                #pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c1 = pv_c1 + "<" + pa_xao_tron + ">"
                                dv_c1 = dv_c1 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c2 = pv_c2 + "<" + pa_xao_tron + ">"
                                dv_c2 = dv_c2 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c3 = pv_c3 + "<" + pa_xao_tron + ">"
                                dv_c3 = dv_c3 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c4 = pv_c4 + "<" + pa_xao_tron + ">"
                                dv_c4 = dv_c4 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c5 = pv_c5 + "<" + pa_xao_tron + ">"
                                dv_c5 = dv_c5 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c6 = pv_c6 + "<" + pa_xao_tron + ">"
                                dv_c6 = dv_c6 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c7 = pv_c7 + "<" + pa_xao_tron + ">"
                                dv_c7 = dv_c7 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c8 = pv_c8 + "<" + pa_xao_tron + ">"
                                dv_c8 = dv_c8 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                                pa_xao_tron, dap_an_xao_tron = xao_tron_de_thi(solgpa, dap_an_pa)
                                pv_c9 = pv_c9 + "<" + pa_xao_tron + ">"
                                dv_c9 = dv_c9 + dap_an_padau[jj:jj+1] + dap_an_xao_tron
                            jj = jj + 2
                        #Sau khi nạp vào hoán vị các phần từ, gán vào 9 h vi
                        skq_dechuan[0]['dapanhv1'] = skq_dechuan[0]['dapanhv1'] + "," + dv_c1
                        skq_dechuan[0]['pahv1'] = skq_dechuan[0]['pahv1'] + "," + pv_c1
                        skq_dechuan[0]['dapanhv2'] = skq_dechuan[0]['dapanhv2'] + "," + dv_c2
                        skq_dechuan[0]['pahv2'] = skq_dechuan[0]['pahv2'] + "," + pv_c2
                        skq_dechuan[0]['dapanhv3'] = skq_dechuan[0]['dapanhv3'] + "," + dv_c3
                        skq_dechuan[0]['pahv3'] = skq_dechuan[0]['pahv3'] + "," + pv_c3
                        skq_dechuan[0]['dapanhv4'] = skq_dechuan[0]['dapanhv4'] + "," + dv_c4
                        skq_dechuan[0]['pahv4'] = skq_dechuan[0]['pahv4'] + "," + pv_c4
                        skq_dechuan[0]['dapanhv5'] = skq_dechuan[0]['dapanhv5'] + "," + dv_c5
                        skq_dechuan[0]['pahv5'] = skq_dechuan[0]['pahv5'] + "," + pv_c5
                        skq_dechuan[0]['dapanhv6'] = skq_dechuan[0]['dapanhv6'] + "," + dv_c6
                        skq_dechuan[0]['pahv6'] = skq_dechuan[0]['pahv6'] + "," + pv_c6
                        skq_dechuan[0]['dapanhv7'] = skq_dechuan[0]['dapanhv7'] + "," + dv_c7
                        skq_dechuan[0]['pahv7'] = skq_dechuan[0]['pahv7'] + "," + pv_c7
                        skq_dechuan[0]['dapanhv8'] = skq_dechuan[0]['dapanhv8'] + "," + dv_c8
                        skq_dechuan[0]['pahv8'] = skq_dechuan[0]['pahv8'] + "," + pv_c8
                        skq_dechuan[0]['dapanhv9'] = skq_dechuan[0]['dapanhv9'] + "," + dv_c9
                        skq_dechuan[0]['pahv9'] = skq_dechuan[0]['pahv9'] + "," + pv_c9
            #print(skq_dechuan[0]['noidungchuan'])
    #print(skq_dechuan[0]['noidungchuan'])
    return skq_dechuan

def xaotroncaccau(skq_dechuan):
    #Hàm này xáo trộn thứ tự các câu
    #print(skq_dechuan[0]['noidungchuan'])
    listnoidunghv = []
    listdapanhv = []
    listpahv = []
    pahvcau = []
    for n in range(1,10):
        noidunghva = []
        dapanhva = []
        pahva = []
        hvcaua = []
        s = skq_dechuan[0][f"noidunghv{n}"]
        vtphay = s.find(",")
        while vtphay >= 0:
            scuoi = s[1:]
            vtphsau = scuoi.find(",")
            if vtphsau < 0:
                st = scuoi
                s = ""
            else:
                st = s[1:vtphsau+1]
                s = s[vtphsau+1:]
            noidunghva.append(st)
            vtphay = s.find(",")
        #đến đáp án hv
        s = skq_dechuan[0][f"dapanhv{n}"]
        vtphay = s.find(",")
        while vtphay >= 0:
            scuoi = s[1:]
            vtphsau = scuoi.find(",")
            if vtphsau < 0:
                st = scuoi
                s = ""
            else:
                st = s[1:vtphsau+1]
                s = s[vtphsau+1:]
            dapanhva.append(st)
            vtphay = s.find(",")
        #đến phương án hoán vị
        s = skq_dechuan[0][f"pahv{n}"]
        vtphay = s.find(",")
        while vtphay >= 0:
            scuoi = s[1:]
            vtphsau = scuoi.find(",")
            if vtphsau < 0:
                st = scuoi
                s = ""
            else:
                st = s[1:vtphsau+1]
                s = s[vtphsau+1:]
            pahva.append(st)
            vtphay = s.find(",")
        #print(dapanhva)
        #Từ đầu đến cuối lấy mã câu hỏi, kiểm tra permuteques = True, dò tìm các câu tiếp theo
        #cùng loại, cùng permuteques = True, đến hết loại.
        slg = len(noidunghva)
        m = 0
        caccauhv_tg = []
        caccauhv_tg_vt = []
        while m < slg:
            caccauhv = []
            loai_hv = models.Subjects001.objects.filter(macauhoi = noidunghva[m]).values('madvkt', 'permuteques')
            if loai_hv[0]['permuteques']:
                ghi_vt = m
                ghi_madvkt = loai_hv[0]['madvkt']
                caccauhv = [m]
                m += 1
                if m >= slg:
                    break
                loai_hv = models.Subjects001.objects.filter(macauhoi = noidunghva[m]).values('madvkt', 'permuteques')
                while loai_hv[0]['permuteques'] and (loai_hv[0]['madvkt'] == ghi_madvkt):
                    caccauhv.append(m)
                    m += 1
                    if m >= slg:
                        break
                    loai_hv = models.Subjects001.objects.filter(macauhoi = noidunghva[m]).values('madvkt', 'permuteques')
            m += 1
            if len(caccauhv) > 5 :
                caccauhv_tg.append(caccauhv)
                caccauhv_tg_vt.append(ghi_vt)
        #Random các câu; caccauhv_tg sẽ có dạng [[1,2,3,4],[5,6,7,8]]
        hv_caccauhv_tg = []
        for i in range(len(caccauhv_tg_vt)):
            tu = caccauhv_tg_vt[i]
            den = caccauhv_tg[i][len(caccauhv_tg[i])-1]
            #---h vi
            hv_caccauhv_tg1 = random.sample(range(tu, den + 1), den - tu + 1)
            hv_caccauhv_tg.append(hv_caccauhv_tg1)
        #print(hv_caccauhv_tg[0])
        #print(skq_dechuan[0]['noidunghv1'])
        #Ghép vào sskq_dechuan1
        noidunghv1s = ""
        dapanhv1s = ""
        pahv1s = ""
        hvcau1s = ""
        j = 0
        slg_doan_hv = 0
        while j < slg:
            if caccauhv_tg_vt[slg_doan_hv] == j :
                for x1 in range(len(hv_caccauhv_tg[slg_doan_hv])):
                    noidunghv1s = noidunghv1s + "," + noidunghva[hv_caccauhv_tg[slg_doan_hv][x1]]
                    dapanhv1s = dapanhv1s + "," + dapanhva[hv_caccauhv_tg[slg_doan_hv][x1]]
                    pahv1s = pahv1s + "," + pahva[hv_caccauhv_tg[slg_doan_hv][x1]]
                    hvcau1s = hvcau1s + "," + str(hv_caccauhv_tg[slg_doan_hv][x1])
                    j += 1
            else:
                noidunghv1s = noidunghv1s + "," + noidunghva[j]
                dapanhv1s = dapanhv1s + "," + dapanhva[j]
                pahv1s = pahv1s + "," + pahva[j]
                hvcau1s = hvcau1s + "," + str(j)
                j += 1
        #-----------------------------
        listnoidunghv.append(noidunghv1s)
        listdapanhv.append(dapanhv1s)
        listpahv.append(pahv1s)
        pahvcau.append(hvcau1s)

    sskq_dechuan = [{"madechuan": skq_dechuan[0]["madechuan"], "fmamon": skq_dechuan[0]["fmamon"], "fmakt": skq_dechuan[0]["fmakt"],
                    "noidungchuan": skq_dechuan[0]["noidungchuan"], "dapanchuan":skq_dechuan[0]["dapanchuan"],
                    "mahoanvi1": skq_dechuan[0]["mahoanvi1"], "noidunghv1": listnoidunghv[0], "dapanhv1": listdapanhv[0], "pahv1": listpahv[0],
                    "mahoanvi2": skq_dechuan[0]["mahoanvi2"], "noidunghv2": listnoidunghv[1], "dapanhv2": listdapanhv[1], "pahv2": listpahv[1],
                    "mahoanvi3": skq_dechuan[0]["mahoanvi3"], "noidunghv3": listnoidunghv[2], "dapanhv3": listdapanhv[2], "pahv3": listpahv[2],
                    "mahoanvi4": skq_dechuan[0]["mahoanvi4"], "noidunghv4": listnoidunghv[3], "dapanhv4": listdapanhv[3], "pahv4": listpahv[3],
                    "mahoanvi5": skq_dechuan[0]["mahoanvi5"], "noidunghv5": listnoidunghv[4], "dapanhv5": listdapanhv[4], "pahv5": listpahv[4],
                    "mahoanvi6": skq_dechuan[0]["mahoanvi6"], "noidunghv6": listnoidunghv[5], "dapanhv6": listdapanhv[5], "pahv6": listpahv[5],
                    "mahoanvi7": skq_dechuan[0]["mahoanvi7"], "noidunghv7": listnoidunghv[6], "dapanhv7": listdapanhv[6], "pahv7": listpahv[6],
                    "mahoanvi8": skq_dechuan[0]["mahoanvi8"], "noidunghv8": listnoidunghv[7], "dapanhv8": listdapanhv[7], "pahv8": listpahv[7],
                    "mahoanvi9": skq_dechuan[0]["mahoanvi9"], "noidunghv9": listnoidunghv[8], "dapanhv9": listdapanhv[8], "pahv9": listpahv[8],
                    "pahvcau1": pahvcau[0], "pahvcau2": pahvcau[1], "pahvcau3": pahvcau[2], "pahvcau4": pahvcau[3], "pahvcau5": pahvcau[4],
                    "pahvcau6": pahvcau[5], "pahvcau7": pahvcau[6], "pahvcau8": pahvcau[7], "pahvcau9": pahvcau[8]}]
    #print(sskq_dechuan)

    return sskq_dechuan

def mkquestions(request):
    """Tạo ngân hàng câu hỏi."""
    assert isinstance(request, HttpRequest)
    vit = models.Dvktm.objects.count()
    
    if vit > 0:
        #dsmon = models.Monhoc.objects.all().values()
        message2r = models.Dvktm.objects.all()[0].tenmon
        mesmamon = models.Dvktm.objects.all()[0].fk_mamon
    else:
        message2r = ""
        mesmamon = ""
    
    tieude_list = []
    # Danh sách combo tiêu đề (giá trị lưu trữ, tên hiển thị)
    if mesmamon != "" :
        kq_mon = models.Monhoc.objects.get(mamon=mesmamon)
        if kq_mon.matd1 != "":
            tieude_list.append([kq_mon.matd1, kq_mon.matd1 + " - " + kq_mon.tentd1])
            if kq_mon.matd2 != "":
                tieude_list.append([kq_mon.matd2, kq_mon.matd2 + " - " + kq_mon.tentd2])
                if kq_mon.madtd3 != "":
                    tieude_list.append([kq_mon.matd3, kq_mon.matd3 + " - " + kq_mon.tentd3])
                    if kq_mon.matd4 != "":
                        tieude_list.append([kq_mon.matd4, kq_mon.matd4 + " - " + kq_mon.tentd4])
                        if kq_mon.matd5 != "":
                            tieude_list.append([kq_mon.matd5, kq_mon.matd5 + " - " + kq_mon.tentd5])

    #Hiển thị ngân hàng câu hỏi
    ch_list = []
    if mesmamon != "" :
        kq_ch = models.Subjects001.objects.filter(mamon=mesmamon)
        ct = 0
        if kq_ch != None :
            for item in kq_ch :
                ct += 1
                #print(item.macauhoi)
                ch_list.append({'tt' : ct, 'nd' : item.questions, 'opt' : item.answer, 'da' : item.correctans, 'dvkt' : item.madvkt, 'td' : item.titles, 'mdo' : item.levelques})
        sttcauhoi = ct + 1
    tongsoch = len(ch_list)
    context ={}
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        hcontents_mucdo = request.POST.get('mucdo')
        hcontents_loaicau = request.POST.get('chonloaicau')
        #textcontent = request.POST.get('content')
        #duong_danr = request.POST.get('ndfile')
        #uploaded_file = request.FILES.get('tenfile')
        if pressed_button == 'save2':
            context['b_form'] = forms.FileCHForm(request.POST, request.FILES)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['b_form'].is_valid():
                #Kiểm tra file có được chọn hay không
                if 'hfile' in request.FILES:
                    # Có file được gửi lên
                    sfile_path = request.FILES['hfile']
                    context['e_form'] = forms.XuatdvktmForm(request.POST)
                    # kiểm tra tính hợp lệ của biểu mẫu:
                    if context['e_form'].is_valid():
                        madvktcombo = context['e_form'].cleaned_data['mavtenm']
                        # Từ mavtendvktr xuất max ddvkt va ten đơn vị kiến thức của môn học
                        if madvktcombo != None :
                            mavtendvktr = madvktcombo.mavtendvkt
                            vi_tr = mavtendvktr.find(" - ")
                            if vi_tr > 0:
                                hcontents_dvkt = mavtendvktr[:vi_tr]
                                ten_dvkt = mavtendvktr[vi_tr + 2 :]
                            else:
                                hcontents_dvkt = ""
                                ten_dvkt = ""
                        else:
                            hcontents_dvkt = ""
                            ten_dvkt = ""
                        if hcontents_dvkt == "00" :
                            hcontents_dvkt = ""
                            ten_dvkt = ""
                        hcontents_mucdo2 = request.POST.get('mucdo2')
                        if hcontents_mucdo2 == "" :
                            hcontents_mucdo2 = "RD"
                        hcontents_tde2 = request.POST.get('tieude_field2')
                        # ------------------------------
                        imp_file(sfile_path, mesmamon, sttcauhoi, ten_dvkt, hcontents_dvkt, hcontents_tde2, hcontents_mucdo2)
                # chuyển hướng đến một URL mới:
                return HttpResponseRedirect('/mkquestions/')
        #elif pressed_button == 'matrancd':

    else:
        context = {
            'message':'Chọn môn học để xây dựng ngân hàng câu hỏi. Lưu ý: Phải chọn môn học trước khi thực hiện các bước tiếp theo.',
            'message2':message2r,
            'nghgch': ch_list,
            'd_form': forms.XuatdvktmForm(),
            'e_form': forms.XuatdvktmForm(),
            'c_form': forms.CauhoiForm(),
            'b_form': forms.FileCHForm,
            'tieudes': tieude_list,
            'tongsoch': tongsoch,
        }

    return render(request, 'app/mkquestions.html', context)

def mkquestions_matrix(request):
    """Tạo ma trận chọn đề."""
    assert isinstance(request, HttpRequest)
    vit = models.Dvktm.objects.count()
    
    if vit > 0:
        #dsmon = models.Monhoc.objects.all().values()
        message2r = models.Dvktm.objects.all()[0].tenmon
        mesmamon = models.Dvktm.objects.all()[0].fk_mamon
    else:
        message2r = ""
        mesmamon = ""
    
    dvkt_list = []
    # Danh sách các đơn vị kiến thức của môn
    if mesmamon != "" :
        kq_dvkt = models.Dvktm.objects.all()
        ct = 0
        if kq_dvkt != None :
            for item in kq_dvkt :
                ma = item.mavtendvkt
                vima = ma.find(" - ")
                if vima > 0:
                    mdvkt = ma[:vima]
                    ct += 1
                    dvkt_list.append({'tt' : ct, 'mdvkt' : mdvkt})
                else:
                    mdvkt = ""
    #Hiển thị ma trận chọn đề đã có
    mtr_list = []
    if mesmamon != "" :
        kq_ch = models.Matranchonde.objects.filter(fk_mamon =mesmamon)
        ct = 0
        if kq_ch != None :
            for item in kq_ch :
                ct += 1
                mtr_list.append({'tt' : ct, 'mamtr' : item.mamtr, 'socaumtr' : item.socaumtr, 'noidungmtr' : item.noidungmtr})
    tongsomtr = len(mtr_list)
    context ={}
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        #hcontents_mucdo = request.POST.get('mucdo')
        #hcontents_loaicau = request.POST.get('chonloaicau')
        #textcontent = request.POST.get('content')
        #duong_danr = request.POST.get('ndfile')
        #uploaded_file = request.FILES.get('tenfile')
        if pressed_button == 'save':
            # kiểm tra tính hợp lệ của biểu mẫu:
            ftenmtr = request.POST.get('tenmtran')
            fttmamtr = tongsomtr + 1
            #mã ma trận đề = mã môn + số thứ tự matrận của môn
            fmamtr = mesmamon + f"{fttmamtr:02d}"
            #Các biến khởi tạo lấy từ đvkt
            fmadvkt = []
            #ftendvkt = ["Không có"]
            if len(dvkt_list) != 0 :
                for ilst in dvkt_list :
                    fmadvkt.append(ilst['mdvkt'])
            ftotalcau = 0
            fthu = ""
            #------lấy dữ liệu các mức độ
            urde = request.POST.getlist('mcdord[]')
            ude = request.POST.getlist('mcdod[]')
            utb = request.POST.getlist('mcdotb[]')
            ukh = request.POST.getlist('mcdokh[]')
            urk = request.POST.getlist('mcdork[]')
            #uname = ftendvkt[0]
            ftenmtr.strip()
            if ftenmtr != "":
                for itm in range(len(fmadvkt)) :
                    #lưu vào csdl
                    ftgsocau = 0
                    if urde[itm] == "" :
                        nurde = "0"
                    else:
                        nurde = urde[itm]
                        ftgsocau = ftgsocau + int(nurde)
                    if ude[itm] == "" :
                        nude = "0"
                    else:
                        nude = ude[itm]
                        ftgsocau = ftgsocau + int(nude)
                    if utb[itm] == "" :
                        nutb = "0"
                    else:
                        nutb = utb[itm]
                        ftgsocau = ftgsocau + int(nutb)
                    if ukh[itm] == "" :
                        nukh = "0"
                    else:
                        nukh = ukh[itm]
                        ftgsocau = ftgsocau + int(nukh)
                    if urk[itm] == "" :
                        nurk = "0"
                    else:
                        nurk = urk[itm]
                        ftgsocau = ftgsocau + int(nurk)
                    if ftgsocau > 0:
                        if fthu == "" :
                            fthu = str(ftgsocau) + "!" + fmadvkt[itm] + "!" + nurde + "!" + nude + "!" + nutb + "!" + nukh + "!" + nurk
                        else:
                            fthu = fthu + "<##>" + str(ftgsocau) + "!" + fmadvkt[itm] + "!" + nurde + "!" + nude + "!" + nutb + "!" + nukh + "!" + nurk
                        ftotalcau = ftotalcau + ftgsocau
                
            #lưu dòng tiếp theo cho đến hết số đvkt
            # Thử lưu vào theo cấu trúc: số câu!mãdvkt!rd!de!tb!kh!rk<##>số câu!mãdvkt!rd!de!tb!kh!rk<##>số câu!mãdvkt!rd!de!tb!kh!rk
            #ftotalcau = 40
            if fthu != "":
                #models.Matranchonde.objects.create(mamtr = fmamtr, fk_mamon = mesmamon, noidungmtr = fthu, socaumtr = ftotalcau)
                print(fthu)
            # chuyển hướng đến một URL mới:
            return HttpResponseRedirect('/mkquestions_matrix/')
        #elif pressed_button == 'matrancd':

    else:
        context = {
            'message':'Chọn môn học để xây dựng ma trận chọn đề. Lưu ý: Phải chọn môn học trước khi thực hiện các bước tiếp theo.',
            'message2':message2r,
            'mtrlst': mtr_list,
            'dvktlst': dvkt_list,
            'tsomtr': tongsomtr,
        }

    return render(request, 'app/mkquestions_matrix.html', context)

def subjectchoice(request):
    """Chọn môn học."""
    assert isinstance(request, HttpRequest)
    context1 ={}
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'chon':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context1['b_form'] = forms.XuatCombomonForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context1['b_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                mamoncombo = context1['b_form'].cleaned_data['hmavten']
                mamonr = mamoncombo.mamon
                # Từ mã môn xuất đơn vị kiến thức của môn học
                if mamonr != None :
                    kq_mon = models.Monhoc.objects.get(mamon=mamonr)
                    ktaodsdvktm = models.Dvktm.objects.count()
                    if ktaodsdvktm > 0:
                        models.Dvktm.objects.all().delete()
                    models.Dvktm.objects.create(mavtendvkt = "00" + " - " + "Không phân chia đvkt/ phần/ chương", fk_mamon = kq_mon.mamon, tenmon = kq_mon.tenmon)
                    if kq_mon.madvkt1 != "":
                        models.Dvktm.objects.create(mavtendvkt = kq_mon.madvkt1 + " - " + kq_mon.tendvkt1, fk_mamon = kq_mon.mamon, tenmon = kq_mon.tenmon)
                        if kq_mon.madvkt2 != "":
                            models.Dvktm.objects.create(mavtendvkt = kq_mon.madvkt2 + " - " + kq_mon.tendvkt2, fk_mamon = kq_mon.mamon, tenmon = kq_mon.tenmon)
                            if kq_mon.madvkt3 != "":
                                models.Dvktm.objects.create(mavtendvkt = kq_mon.madvkt3 + " - " + kq_mon.tendvkt3, fk_mamon = kq_mon.mamon, tenmon = kq_mon.tenmon)
                                if kq_mon.madvkt4 != "":
                                    models.Dvktm.objects.create(mavtendvkt = kq_mon.madvkt4 + " - " + kq_mon.tendvkt4, fk_mamon = kq_mon.mamon, tenmon = kq_mon.tenmon)
                                    if kq_mon.madvkt5 != "":
                                        models.Dvktm.objects.create(mavtendvkt = kq_mon.madvkt5 + " - " + kq_mon.tendvkt5, fk_mamon = kq_mon.mamon, tenmon = kq_mon.tenmon)
                
                    # chuyển hướng đến một URL mới:
                    return HttpResponseRedirect('/mkquestions/')

    else:
        context1 = {
            'message':'Trang này chọn môn học để xây dựng ngân hàng câu hỏi.',
            'b_form': forms.XuatCombomonForm(),
            #'h_x': Usercs.objects.get(id=1),
        }

    return render(request, 'app/subjectchoice.html', context1)


def saveunitspart():
    #x = "Subjects" + "001"
    #tenbangmon = models.x
    return "Giá trị trả về"

def unitspart(request):
    # Thêm môn học mới và đơn vị kiến thức, tiêu đề
    vitri = models.Monhoc.objects.count()
    #if vitri > 0:
    #    danhsachmon = models.Monhoc.objects.all().values()
    #----------------
    context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context['h_form'] = forms.MonhocForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['h_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                ma = context['h_form'].cleaned_data['hmamon']
                ten = context['h_form'].cleaned_data['htenmon']
                mng = context['h_form'].cleaned_data['hfk_manganh']
                trd = context['h_form'].cleaned_data['htrinhdo']
                sgi = context['h_form'].cleaned_data['hsogio']
                mdv1 = context['h_form'].cleaned_data['hmadvkt1']
                tdv1 = context['h_form'].cleaned_data['htendvkt1']
                mdv2 = context['h_form'].cleaned_data['hmadvkt2']
                tdv2 = context['h_form'].cleaned_data['htendvkt2']
                mdv3 = context['h_form'].cleaned_data['hmadvkt3']
                tdv3 = context['h_form'].cleaned_data['htendvkt3']
                mdv4 = context['h_form'].cleaned_data['hmadvkt4']
                tdv4 = context['h_form'].cleaned_data['htendvkt4']
                mdv5 = context['h_form'].cleaned_data['hmadvkt5']
                tdv5 = context['h_form'].cleaned_data['htendvkt5']
                mtd1 = context['h_form'].cleaned_data['hmatd1']
                ttd1 = context['h_form'].cleaned_data['htentd1']
                mtd2 = context['h_form'].cleaned_data['hmatd2']
                ttd2 = context['h_form'].cleaned_data['htentd2']
                mtd3 = context['h_form'].cleaned_data['hmatd3']
                ttd3 = context['h_form'].cleaned_data['htentd3']
                mtd4 = context['h_form'].cleaned_data['hmatd4']
                ttd4 = context['h_form'].cleaned_data['htentd4']
                mtd5 = context['h_form'].cleaned_data['hmatd5']
                ttd5 = context['h_form'].cleaned_data['htentd5']
                #vitri = models.Monhoc.objects.count()
                #vi_tri_ngoac1 = ten.find("(") + 1
                #vi_tri_ngoac2 = ten.find(")") - 1
                if mng != "":
                    mng2 = mng.manganh
                    ma = mng2 + ma
                else:
                    mng2 = ""
                thu = [mdv1, mdv2, mdv3, mdv4, mdv5]
                vt = 0
                for item in thu:
                    if item != "":
                        vt = vt + 1
                    else:
                        break
                if vt > 0:
                    thu2 = thu[0:vt-1]
                    khongtrung = len(thu2) == len(set(thu2))
                else:
                    khongtrung = True
                # cho tiêu đề
                thu3 = [mtd1, mtd2, mtd3, mtd4, mtd5]
                vt = 0
                for item in thu3:
                    if item != "":
                        vt = vt + 1
                    else:
                        break
                if vt > 0:
                    thu4 = thu3[0:vt-1]
                    khongtrung1 = len(thu4) == len(set(thu4))
                else:
                    khongtrung1 = True
                # -------------------
                if vitri > 0 and khongtrung and khongtrung1:
                    if models.Monhoc.objects.filter(mamon=ma).exists():
                        khong = True
                    else:
                        if ma != "":
                            models.Monhoc.objects.create(mamon=ma, tenmon=ten, fk_manganh=mng2, trinhdo=trd, sogio=sgi,
                                                         madvkt1=mdv1, tendvkt1=tdv1, madvkt2=mdv2, tendvkt2=tdv2,
                                                         madvkt3=mdv3, tendvkt3=tdv3, madvkt4=mdv4, tendvkt4=tdv4,
                                                         madvkt5=mdv5, tendvkt5=tdv5, madvkt6="", tendvkt6="",
                                                         madvkt7="", tendvkt7="", madvkt8="", tendvkt8="",madvkt9="", tendvkt9="",
                                                         madvkt10="", tendvkt10="", matd1=mtd1, tentd1=ttd1, matd2=mtd2, tentd2=ttd2,
                                                         matd3=mtd3, tentd3=ttd3, matd4=mtd4, tentd4=ttd4, matd5=mtd5, tentd5=ttd5)
                    
                elif ma != "":
                    models.Monhoc.objects.create(mamon=ma, tenmon=ten, fk_manganh=mng2, trinhdo=trd, sogio=sgi,
                                                 madvkt1=mdv1, tendvkt1=tdv1, madvkt2=mdv2, tendvkt2=tdv2,
                                                 madvkt3=mdv3, tendvkt3=tdv3, madvkt4=mdv4, tendvkt4=tdv4,
                                                 madvkt5=mdv5, tendvkt5=tdv5, madvkt6="", tendvkt6="",
                                                 madvkt7="", tendvkt7="", madvkt8="", tendvkt8="",madvkt9="", tendvkt9="",
                                                 madvkt10="", tendvkt10="", matd1=mtd1, tentd1=ttd1, matd2=mtd2, tentd2=ttd2,
                                                 matd3=mtd3, tentd3=ttd3, matd4=mtd4, tentd4=ttd4, matd5=mtd5, tentd5=ttd5)
                
                # chuyển hướng đến một URL mới:
                return HttpResponseRedirect('/unitspart/')
        elif pressed_button == 'delete':
            # Xử lý logic xóa dữ liệu ở đây
            return HttpResponseRedirect('/about/')
        elif pressed_button == 'cancel':
            # Xử lý logic khi hủy bỏ
            return HttpResponseRedirect('/about/')

    else:
        context['h_form']= forms.MonhocForm()

    return render(request, 'app/unitspart.html', context)
    
def unitspartviews(request):
    """Hiện bảng môn học."""
    assert isinstance(request, HttpRequest)
    vitri = models.Monhoc.objects.count()
    context = {}
    if vitri > 0:
        context = models.Monhoc.objects.all().values()
        #ohde = '<tr><td>' + '1' + '<td>' + danhsachnganh['hmanganh'] + '</td><td>' + danhsachnganh['htennganh'] + '</td><td>' + danhsachnganh['htrinhdo'] + '</td><td>' + danhsachnganh['hkhoa'] + '</td></tr>'
    
    return render(request, 'app/unitspartviews.html', {'data': context})

def unitspartrepair(request):
    # Sửa đơn vị kiến thức
    vitri = models.Monhoc.objects.count()
    if vitri > 0:
        danhsachmon = models.Monhoc.objects.all().values()
    #----------------
    context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context['m_form'] = forms.MonhocFormrep(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['m_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                matrc = context['m_form'].cleaned_data['hmamoned1']
                matrc = matrc.strip()
                co = False
                # kiểm tra mã có trong danh sách hay chưa
                if danhsachmon[0].get('mamon') == matrc:
                    co = True
                elif danhsachmon[len(danhsachmon)-1].get('mamon') == matrc:
                    co = True
                else:
                    for row in range(len(danhsachmon)):
                        if matrc == danhsachmon[row].get('mamon'):
                            co = True
                            break
                if co:
                    mas1 = context['m_form'].cleaned_data['hmamoned']
                    tens = context['m_form'].cleaned_data['htenmoned']
                    mngs = context['m_form'].cleaned_data['hfk_manganhed']
                    trds = context['m_form'].cleaned_data['htrinhdoed']
                    sgis = context['m_form'].cleaned_data['hsogioed']
                    mdv1s = context['m_form'].cleaned_data['hmadvkt1ed']
                    tdv1s = context['m_form'].cleaned_data['htendvkt1ed']
                    mdv2s = context['m_form'].cleaned_data['hmadvkt2ed']
                    tdv2s = context['m_form'].cleaned_data['htendvkt2ed']
                    mdv3s = context['m_form'].cleaned_data['hmadvkt3ed']
                    tdv3s = context['m_form'].cleaned_data['htendvkt3ed']
                    mdv4s = context['m_form'].cleaned_data['hmadvkt4ed']
                    tdv4s = context['m_form'].cleaned_data['htendvkt4ed']
                    mdv5s = context['m_form'].cleaned_data['hmadvkt5ed']
                    tdv5s = context['m_form'].cleaned_data['htendvkt5ed']
                    mtd1 = context['m_form'].cleaned_data['hmatd1ed']
                    ttd1 = context['m_form'].cleaned_data['htentd1ed']
                    mtd2 = context['m_form'].cleaned_data['hmatd2ed']
                    ttd2 = context['m_form'].cleaned_data['htentd2ed']
                    mtd3 = context['m_form'].cleaned_data['hmatd3ed']
                    ttd3 = context['m_form'].cleaned_data['htentd3ed']
                    mtd4 = context['m_form'].cleaned_data['hmatd4ed']
                    ttd4 = context['m_form'].cleaned_data['htentd4ed']
                    mtd5 = context['m_form'].cleaned_data['hmatd5ed']
                    ttd5 = context['m_form'].cleaned_data['htentd5ed']
                    #vitri = models.Monhoc.objects.count()
                    #vi_tri_ngoac1 = ten.find("(") + 1
                    #vi_tri_ngoac2 = ten.find(")") - 1
                    if mngs != "":
                        mng2 = mngs.manganh
                        mas = mng2 + mas1
                    else:
                        mng2 = ""
                    thu = [mdv1s, mdv2s, mdv3s, mdv4s, mdv5s]
                    vt = 0
                    for item in thu:
                        if item != "":
                            vt = vt + 1
                        else:
                            break
                    if vt > 0:
                        thu2 = thu[0:vt-1]
                        khongtrung = len(thu2) == len(set(thu2))
                    else:
                        khongtrung = True
                    # cho tiêu đề
                    thu3 = [mtd1, mtd2, mtd3, mtd4, mtd5]
                    vt = 0
                    for item in thu3:
                        if item != "":
                            vt = vt + 1
                        else:
                            break
                    if vt > 0:
                        thu4 = thu3[0:vt-1]
                        khongtrung1 = len(thu4) == len(set(thu4))
                    else:
                        khongtrung1 = True
                    # -------------------
                    
                    if vitri > 0 and khongtrung and khongtrung1:
                        if models.Monhoc.objects.filter(mamon=mas).exists():
                            khong = True
                        else:
                            if mas != "":
                                models.Monhoc.objects.filter(mamon = matrc).update(mamon=mas, tenmon=tens, fk_manganh=mng2, trinhdo=trds, sogio=sgis,
                                                         madvkt1=mdv1s, tendvkt1=tdv1s, madvkt2=mdv2s, tendvkt2=tdv2s,
                                                         madvkt3=mdv3s, tendvkt3=tdv3s, madvkt4=mdv4s, tendvkt4=tdv4s,
                                                         madvkt5=mdv5s, tendvkt5=tdv5s, madvkt6="", tendvkt6="",
                                                         madvkt7="", tendvkt7="", madvkt8="", tendvkt8="",madvkt9="", tendvkt9="",
                                                         madvkt10="", tendvkt10="", matd1=mtd1, tentd1=ttd1, matd2=mtd2, tentd2=ttd2,
                                                         matd3=mtd3, tentd3=ttd3, matd4=mtd4, tentd4=ttd4, matd5=mtd5, tentd5=ttd5)
                    
                    #Dùng phương thức save sẽ thêm vào ở cuối csdl, nên dùng update()
                # chuyển hướng đến một URL mới:
                return HttpResponseRedirect('/unitspartrepair/')
        elif pressed_button == 'delete':
            # Xử lý logic xóa dữ liệu ở đây
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context['m_form'] = forms.MonhocFormrep(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['m_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                matrc = context['m_form'].cleaned_data['hmamoned1']
                matrc = matrc.strip()
                co = False
                # kiểm tra mã có trong danh sách hay chưa
                if danhsachmon[0].get('mamon') == matrc:
                    co = True
                elif danhsachmon[len(danhsachmon)-1].get('mamon') == matrc:
                    co = True
                else:
                    for row in range(len(danhsachmon)):
                        if matrc == danhsachmon[row].get('mamon'):
                            co = True
                            break
                if co:
                    models.Monhoc.objects.filter(mamon = matrc).delete()

            return HttpResponseRedirect('/about/')
        elif pressed_button == 'cancel':
            # Xử lý logic khi hủy bỏ
            return HttpResponseRedirect('/about/')

    else:
        context['m_form'] = forms.MonhocFormrep()
        
    return render(request, 'app/unitspartrepair.html', context)


def apmajor(request):
    # Tạo ngành học
    context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context['ng_form'] = forms.NganhhocForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['ng_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                ma = context['ng_form'].cleaned_data['hmanganh']
                ten = context['ng_form'].cleaned_data['htennganh']
                trdo = context['ng_form'].cleaned_data['htrinhdo']
                kh = context['ng_form'].cleaned_data['hkhoa']
                vitri = models.Nganhhoc.objects.count()
                if vitri > 0:
                    x = models.Nganhhoc.objects.all()[vitri-1]
                if ma != "":
                    models.Nganhhoc.objects.create(manganh = ma, tennganh = ten, trinhdo = trdo, khoa = kh)
    
                # chuyển hướng đến một URL mới:
                return HttpResponseRedirect('/apmajor/')
        elif pressed_button == 'delete':
            # Xử lý logic xóa dữ liệu ở đây
            return HttpResponseRedirect('/about/')
        elif pressed_button == 'cancel':
            # Xử lý logic khi hủy bỏ
            return HttpResponseRedirect('/about/')

    else:
        context['ng_form'] = forms.NganhhocForm()
        
    return render(request, 'app/apmajor.html', context)

def apmajorrepair(request):
    # Sửa chữa ngành học
    vitri = models.Nganhhoc.objects.count()
    if vitri > 0:
        danhsachnganh = models.Nganhhoc.objects.all().values()
    #----------------
    context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context['ng_form'] = forms.NganhhocFormrep(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['ng_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                matrc = context['ng_form'].cleaned_data['hmanganhed1']
                matrc = matrc.strip()
                co = False
                # kiểm tra mã có trong danh sách hay chưa
                if danhsachnganh[0].get('manganh') == matrc:
                    co = True
                    vit = 0
                elif danhsachnganh[len(danhsachnganh)-1].get('manganh') == matrc:
                    co = True
                    vit = len(danhsachnganh)-1
                else:
                    for row in range(len(danhsachnganh)):
                        if matrc == danhsachnganh[row].get('manganh'):
                            co = True
                            vit = row
                            break
                if co:
                    mas = context['ng_form'].cleaned_data['hmanganhed']
                    tens = context['ng_form'].cleaned_data['htennganhed']
                    trdos = context['ng_form'].cleaned_data['htrinhdoed']
                    khs = context['ng_form'].cleaned_data['hkhoaed']
                    models.Nganhhoc.objects.filter(manganh = matrc).update(manganh = mas, tennganh = tens, trinhdo = trdos, khoa = khs)
                    #models.Nganhhoc.objects.filter(manganh = matrc).update(manganh = mas)
                    #models.Nganhhoc.objects.all()[vit].save()
                    #Dùng phương thức save sẽ thêm vào ở cuối csdl, nên dùng update()
                
                # chuyển hướng đến một URL mới:
                return HttpResponseRedirect('/apmajorrepair/')
        elif pressed_button == 'delete':
            # Xử lý logic xóa dữ liệu ở đây
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context['ng_form'] = forms.NganhhocFormrep(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['ng_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                matrc = context['ng_form'].cleaned_data['hmanganhed1']
                matrc = matrc.strip()
                co = False
                # kiểm tra mã có trong danh sách hay chưa
                if danhsachnganh[0].get('manganh') == matrc:
                    co = True
                elif danhsachnganh[len(danhsachnganh)-1].get('manganh') == matrc:
                    co = True
                else:
                    for row in range(len(danhsachnganh)):
                        if matrc == danhsachnganh[row].get('manganh'):
                            co = True
                            break
                if co:
                    models.Nganhhoc.objects.filter(manganh = matrc).delete()

            return HttpResponseRedirect('/about/')
        elif pressed_button == 'cancel':
            # Xử lý logic khi hủy bỏ
            return HttpResponseRedirect('/about/')

    else:
        context['ng_form'] = forms.NganhhocFormrep()
        
    return render(request, 'app/apmajorrepair.html', context)


def apmajorviews(request):
    """Xem ngành học."""
    assert isinstance(request, HttpRequest)
    #danhsachnganh = {'hmanganh': "Đây là mã ngành", 'htennganh': "Đây là tên ngành", 'htrinhdo': "Đây là trình độ", 'hkhoa': "Đây là khoa"}
    vitri = models.Nganhhoc.objects.count()
    if vitri > 0:
        danhsachnganh = models.Nganhhoc.objects.all().values()
        #x = models.Nganhhoc.objects.all()[0]
        #danhsachnganh = {'hmanganh': x.manganh, 'htennganh': x.tennganh, 'htrinhdo': x.trinhdo, 'hkhoa': x.khoa}
        #ohde = '<tr><td>' + '1' + '<td>' + danhsachnganh['hmanganh'] + '</td><td>' + danhsachnganh['htennganh'] + '</td><td>' + danhsachnganh['htrinhdo'] + '</td><td>' + danhsachnganh['hkhoa'] + '</td></tr>'

    #--------------------
    #context = {'TT': 1}
    #context.update(danhsachnganh)
    context = danhsachnganh
    return render(request, 'app/apmajorviews.html', {'data': context})

def settest(request):
    """Đưa các thông tin về kỳ thi, môn thi để lưu vào session."""
    assert isinstance(request, HttpRequest)
    context1 ={}
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'chon':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context1['a_form'] = forms.XuatKythiForm(request.POST)
            context1['b_form'] = forms.XuatCombomonForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context1['a_form'].is_valid() and context1['b_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                maktcombo = context1['a_form'].cleaned_data['hfk_makt']
                maktr = maktcombo.makt
                mamoncombo = context1['b_form'].cleaned_data['hmavten']
                mamonr = mamoncombo.mamon
                # Lưu vào session 2 biến kỳ thi và môn học
                # Lưu giá trị vào session
                if maktr != None :
                    models.CurrentSetting.objects.filter(tt = 1).update(makt = maktr)
                    #request.session['hkythi'] = maktr
                    request.session['hmonthi'] = mamonr
                #return HttpResponseRedirect('/mkquestions/')

    else:
        context1 = {
            'a_form': forms.XuatKythiForm(),
            'b_form': forms.XuatCombomonForm(),
            #'h_x': Usercs.objects.get(id=1),
        }
    return render(request, 'app/settest.html', context1)

def makeexam(request):
    """Tạo kỳ thi."""
    assert isinstance(request, HttpRequest)
    vitri = models.Kythi.objects.count() + 1
    context = {}
    prodkt = models.Kythi.objects.all()
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context['h_form'] = forms.MakeExamForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['h_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                ten = context['h_form'].cleaned_data['htenkt']
                mnh = context['h_form'].cleaned_data['hfk_manh']
                mtkt = context['h_form'].cleaned_data['hmotakt']
                if mnh != "":
                    mng2 = mnh.manh
                    ma = mnh.manh + f"{vitri:03d}"
                else:
                    ma = ""
                    mng2 = ""
                
                if ma != "":
                    if models.Kythi.objects.filter(makt=ma).exists():
                        khong = True
                    else:
                        models.Kythi.objects.create(makt=ma, tenkythi=ten, motakt = mtkt, fkmanh=mng2)
                
                # chuyển hướng đến một URL mới:
                return HttpResponseRedirect('/makeexam/')
        elif pressed_button == 'delete':
            # Xử lý logic xóa dữ liệu ở đây
            context['h_form'] = forms.MakeExamForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['h_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                ma = context['h_form'].cleaned_data['hmakt']
                if ma != "":
                    if models.Kythi.objects.filter(makt=ma).exists():
                        models.Kythi.objects.filter(makt=ma).delete()
                    
            return HttpResponseRedirect('/makeexam/')
        elif pressed_button == 'cancel':
            # Xử lý logic khi hủy bỏ
            return HttpResponseRedirect('/settest/')

    else:
        h_form = forms.MakeExamForm()
        #context['h_form']= forms.MakeExamForm()
    context = {'h_form': h_form, 'prodkt': prodkt}

    return render(request, 'app/makeexam.html', context)

def createstandard(request):
    """Tạo đề chuẩn của kỳ thi."""
    assert isinstance(request, HttpRequest)
    #vitri = models.Kythi.objects.count() + 1
    context = {}
    #prodkt = models.Kythi.objects.all()
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    #Lấy kỳ thi hiện hành
    xds = models.CurrentSetting.objects.filter(tt = 1).values('makt')
    hu_makt = xds[0]['makt']
    #if 'hkythi' in request.session:
    #    hu_makt = request.session.get('hkythi')
    #else:
    #    hu_makt = ""
    if 'hmonthi' in request.session:
        hu_mamon = request.session.get('hmonthi')
    else:
        hu_mamon = ""
    if 'hmatrande' in request.session:
        hu_mamtr = request.session.get('hmatrande')
    else:
        hu_mamtr = ""
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save':
            #Lấy mã kỳ thi, mã môn thi, mã ma trận chọn đề
            if (hu_makt != "" and hu_mamon != "" and hu_mamtr != "") :
                # tạo mã đề chuẩn và các nội dung của đề chuẩn
                # Mã đề chuẩn dạng: mã môn + 0001
                # Kiểm tra thứ tự mã đề chuẩn của môn này theo dạng trên
                gialap_mdchuan = models.Cacdethi.objects.filter(fmamon = hu_mamon)
                slgmdc = len(gialap_mdchuan) + 1
                hu_madechuan = hu_mamon + f"{slgmdc:04d}"
                #tạo list mã đề hoán vị dạng [102,212,395,4..]
                lmahoanvi = []
                for x in range(9):
                    #tạo số ngẫu nhiên 2 chữ số để nối với 1 đến 9
                    x2 = random.sample(range(99), 1)
                    #tạo mã hoán vị
                    lmahoanvi.append(f'{x + 1}{x2[0]:02d}')
                #Lấy ma trận đề chọn
                
                results_mtr = models.Matranchonde.objects.filter(mamtr = hu_mamtr)
                #print(results_mtr)
                listnoidung = []
                ndtach = results_mtr[0].noidungmtr
                vtcach = ndtach.find("<##>")#vị trí phân cách các phần của từng đơn vị kiến thức
                while vtcach >= 0:
                    while vtcach == 0:
                        ndtach = ndtach[4:]
                        vtcach = ndtach.find("<##>")#thêm chỗ này
                    vtthang = ndtach.find("!")
                    socau = int(ndtach[:vtthang])
                    ndtach = ndtach[vtthang+1:]
                    vtthang = ndtach.find("!")
                    noidung = [ndtach[:vtthang]]#lấy mã đvkt
                    ndtach = ndtach[vtthang+1:]
                    vtthang = ndtach.find("!")
                    sc = int(ndtach[:vtthang])
                    noidung.append(sc)# lấy số câu rde
                    ndtach = ndtach[vtthang+1:]
                    vtthang = ndtach.find("!")
                    sc = int(ndtach[:vtthang])
                    noidung.append(sc)# lấy số câu de
                    ndtach = ndtach[vtthang+1:]
                    vtthang = ndtach.find("!")
                    sc = int(ndtach[:vtthang])
                    noidung.append(sc)# lấy số câu tb
                    ndtach = ndtach[vtthang+1:]
                    vtthang = ndtach.find("!")
                    sc = int(ndtach[:vtthang])
                    noidung.append(sc)# lấy số câu kh
                    ndtach = ndtach[vtthang+1:]
                    vtthang = ndtach.find("<##>")
                    sc = int(ndtach[:vtthang])
                    noidung.append(sc)# lấy số câu rk
                    ndtach = ndtach[vtthang+4:]
                    listnoidung.append(noidung)#Nhập vào list tong nếu có 2 đvkt trở lên
                    vtcach = ndtach.find("<##>")
                
                if vtcach < 0:
                    vtthang = ndtach.find("!")
                    socau = int(ndtach[:vtthang])
                    ndtach = ndtach[vtthang+1:]
                    vtthang = ndtach.find("!")
                    noidung = [ndtach[:vtthang]]#lấy mã đvkt
                    ndtach = ndtach[vtthang+1:]
                    vtthang = ndtach.find("!")
                    sc = int(ndtach[:vtthang])
                    noidung.append(sc)# lấy số câu rde
                    ndtach = ndtach[vtthang+1:]
                    vtthang = ndtach.find("!")
                    sc = int(ndtach[:vtthang])
                    noidung.append(sc)# lấy số câu de
                    ndtach = ndtach[vtthang+1:]
                    vtthang = ndtach.find("!")
                    sc = int(ndtach[:vtthang])
                    noidung.append(sc)# lấy số câu tb
                    ndtach = ndtach[vtthang+1:]
                    vtthang = ndtach.find("!")
                    sc = int(ndtach[:vtthang])
                    noidung.append(sc)# lấy số câu kh
                    ndtach = ndtach[vtthang+1:]
                    sc = int(ndtach)
                    noidung.append(sc)# lấy số câu rk
                    #ndtach = ndtach[vtthang+4:]
                    listnoidung.append(noidung)#Nhập vào list tong nếu có 2 đvkt trở lên
                
                #Lấy ngẫu nhiên các câu trong môn học với số lượng theo ma trận đề
                kq_dechuan1 = taodechuanvhoanvi(hu_makt, hu_mamon, listnoidung, hu_madechuan, lmahoanvi)
                #Lưu vào table Cacdethi
                #Xáo trộn các câu hỏi, tạm thời chỉ xáo trộn các câu loại LC và tổng nhóm
                kq_dechuan = xaotroncaccau(kq_dechuan1)
                # Chuyển đổi dict thành instance của Cacdethi trước khi truyền vào bulk_create
                h_instances = [models.Cacdethi(**item) for item in kq_dechuan]
                models.Cacdethi.objects.bulk_create(h_instances)
                #Person.objects.bulk_create(objs, batch_size=1000)
            # chuyển hướng đến một URL mới:
            return HttpResponseRedirect('/createstandard/')
    else:
        context = {
            'x': hu_makt,
            'y': hu_mamon,
            'z': hu_mamtr,
        }

    return render(request, 'app/createstandard.html', context)

def gethoanvi(request):
    """Renders the about page."""
    #Các đề chuẩn sử dụng cho kỳ thi
    assert isinstance(request, HttpRequest)
    #Từ kỳ thi đã thiết lập cho đợt này, hiện bảng các môn thi và mã đề chuẩn của môn để đánh dấu chọn
    #models.Subjects001.objects.all().delete()
    #vitri = models.Namhoc.objects.count()
    xds = models.CurrentSetting.objects.filter(tt = 1).values('makt')
    ukythi = xds[0]['makt']
    #if 'hkythi' in request.session:
    #    ukythi = request.session.get('hkythi')
    #else:
    #    ukythi = ""
    tbldanhsach = []
    if ukythi != "":
        dsmade = models.Cacdethi.objects.filter(fmakt = ukythi).values('madechuan', 'fmamon', 'fmakt')
        #dsmade = models.Cacdethi.objects.all()
        #ra ds gồm mã đề chuẩn, tên môn, lần thi, chọn.
        for d in range(len(dsmade)):
            tenmon1 = models.Monhoc.objects.filter(mamon = dsmade[d]['fmamon']).values('tenmon')
            tenmon = tenmon1[0]['tenmon']
            tbldanhsach.append({'tt' : d + 1, 'ten' : tenmon, 'ma' : dsmade[d]['madechuan'], 'kythi' : dsmade[d]['fmakt'], 'lanthi' : 1})
    #----------------
    context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save':
            # tạo danh sách mới dựa trên check các đề thi được sử dụng:
            tbldanhsach_new = []
            # Lấy danh sách ID của các dòng được tích chọn checkbox (name="selected_ids")
            chonck = request.POST.getlist('selected_ids')
            ich = 0
            for rds in range(len(tbldanhsach)):
                if rds + 1 == int(chonck[ich]):#1,3
                    tbldanhsach_new.append({'madechuan' : dsmade[rds]['madechuan'], 'fmamon' : dsmade[rds]['fmamon'], 'fmakt' : dsmade[rds]['fmakt'], 'lanthi' : tbldanhsach[rds]['lanthi']})
                    ich += 1
                    
            #print(tbldanhsach_new)
            # Tại đây bạn có thể thực hiện logic xử lý tiếp theo (xóa, cập nhật trạng thái...)
            if len(tbldanhsach_new) > 0:
                models.Cacdethikynay.objects.all().delete()
                # Chuyển đổi dict thành instance của Cacdethikynay trước khi truyền vào bulk_create
                h_instances = [models.Cacdethikynay(**item) for item in tbldanhsach_new]
                models.Cacdethikynay.objects.bulk_create(h_instances)
            return HttpResponseRedirect('/gethoanvi/')

    else:
        context = {
            'title': 'Thiết lập các đề thi',
            'year' : datetime.now().year,
            'tbldanhsach': tbldanhsach,
        }
   
    return render(request, 'app/gethoanvi.html', context)

def getmtrde(request):
    """Tạo đề chuẩn của kỳ thi."""
    assert isinstance(request, HttpRequest)
    #vitri = models.Kythi.objects.count() + 1
    context = {}
    if 'hmonthi' in request.session:
        y = request.session.get('hmonthi')
    else:
        y = ""
    prodmtr = models.Matranchonde.objects.filter(fk_mamon = y)
    #prodmtr = models.Matranchonde.objects.all()
    #Dòng này cập nhât thủ công mã ma trận chọn đề
    #models.Matranchonde.objects.filter(fk_mamon = y).update(mamtr = "0168102010701")
    #models.Matranchonde.objects.filter(fk_mamon = y).update(mamtr = "0400000000401")
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context['m_form'] = forms.GetMatrandeForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['m_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                hmamtr = context['m_form'].cleaned_data['hfk_mamtr']
                if hmamtr != "":
                    if models.Matranchonde.objects.filter(mamtr=hmamtr).exists():
                        request.session['hmatrande'] = hmamtr
                        #messages.success(request, "Chọn ma trận chọn đề thành công!")
                
                # chuyển hướng đến một URL mới:
                return HttpResponseRedirect('/createstandard/')
    else:
        m_form = forms.GetMatrandeForm()
        #context['h_form']= forms.MakeExamForm()
    context = {'m_form': m_form, 'prodmtr': prodmtr}

    return render(request, 'app/getmtrde.html', context)

def hien(request):
    """Renders the about page."""
    assert isinstance(request, HttpRequest)
    #danhsachnganh = {'hmanganh': "Đây là mã ngành", 'htennganh': "Đây là tên ngành", 'htrinhdo': "Đây là trình độ", 'hkhoa': "Đây là khoa"}
    #models.Subjects001.objects.create(macauhoi = "01", titles = "con khỉ")
    #models.Usercs.objects.all().delete()
    #vitri = models.Cacdethi.objects.count()
    #models.Cacdethi.objects.filter(madechuan = "016810201250006").delete()
    #models.Lophocphan.objects.filter(malophp = "01681020107001").delete()
    #models.CurrentSetting.objects.create(tt = 1, fkmanh = "2025-2026", makt = "2025-2026002")
    """
    if request.method == 'POST' and request.FILES.get('image_field'):
    # Lấy đối tượng file từ request.FILES
    uploaded_file = request.FILES['image_field']

    # 1. Lấy tên file
    file_name = uploaded_file.name

    # 2. Đọc nội dung file (dạng bytes)
    file_content = uploaded_file.read()
    """
    #models.Lophocphan.objects.filter(malophp = "01681020125001").update(malophp = umalophp, tenlophp = utenlophp, mota = umota, fkmanh = ufkmanh, fkmamon = mahp)
    #print(noidunghva)
    listitems = [
        {'tt' : 1, 'mch' : "nguyeexn van a", 'nd' : "hhha@yahoo.com.vn"},
        {'tt' : 2, 'mch' : "nguyeexn van b", 'nd' : "hhhb@yahoo.com.vn"},
        {'tt' : 3, 'mch' : "nguyeexn van c", 'nd' : "hhhc@yahoo.com.vn"},
        {'tt' : 4, 'mch' : "nguyeexn van d", 'nd' : "hhhd@yahoo.com.vn"}
    ]

    if request.method == 'POST':
        item_id = request.POST.get('item_id')
        name = request.POST.get('name')
        email = request.POST.get('email')

    context = {
        'items': listitems,
    }
    
    #del request.session['ukythi']
    #del request.session['umonthi']
    #request.session.flush()
    #context = {'mot': x, 'hai': y, 'ba': z, 'bon': a, 'nam': b, 'sau': c, 'bay': d, 'tam': e, 'chin': f, 'muoi': g}
    return render(request, 'app/hien.html', context)

def khoi_tao_dl_subj(smadechuan):
    #random để lấy đề số thứ tự hoán vị
    QUIZ_DATA = []
    #random mã đề hoán vị
    sng = random.randint(1, 9)
    #Từ mã đề chuẩn, lấy danh sách các biến
    smahvi = ""
    if smadechuan != "":
        quiz = models.Cacdethi.objects.filter(madechuan = smadechuan).values('fmamon', 'fmakt', 'noidungchuan', 'dapanchuan', f'mahoanvi{sng}', f'noidunghv{sng}', f'dapanhv{sng}', f'pahv{sng}')
    if quiz != "":
        #Khoi tao cac bien, lấy phương án hoán vị, vào Subjects001 lấy nội dung câu hỏi
        sdvkt = ""
        sdvkt_tendung = ""
        stieude = ""
        stieude_tendung = ""
        scauhoichung = ""
        sttcauhoi = 0
        scauhoi = ""
        option = {}
        sdapan = ""
        #assscauhoi = sdvkt + stieude + scauhoichung
        nhan_dap_an = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
        #---------------------
        xmamon = quiz[0]['fmamon']
        macauhoi = quiz[0][f'noidunghv{sng}']
        dapan = quiz[0][f'dapanhv{sng}']
        pahv = quiz[0][f'pahv{sng}']
        smahvi = quiz[0][f'mahoanvi{sng}']
        ipic = ""
        #lấy phương án hoán vị, vào Subjects001 lấy nội dung câu hỏi
        vtphay = macauhoi.find(",")
        while vtphay >= 0:
            macauhoi1 = macauhoi[1:16]
            macauhoi = macauhoi[16:]
            vtphay = macauhoi.find(",")
            fullcauhoi = models.Subjects001.objects.filter(macauhoi = macauhoi1).values()
            #----------
            pahvcuoi = pahv[1:]
            vtphaypahvcuoi = pahvcuoi.find(",")
            if vtphaypahvcuoi >= 0 :
                pahv1 = pahv[1:vtphaypahvcuoi+1]
                pahv = pahv[vtphaypahvcuoi+1:]
            else:
                pahv1 = pahvcuoi
                pahv = ""
            #----------
            dapancuoi = dapan[1:]
            vtphaydapancuoi = dapancuoi.find(",")
            if vtphaydapancuoi < 0:
                dapan1 = dapancuoi
                dapan = ""
            else:
                dapan1 = dapan[1:vtphaydapancuoi+1]
                dapan = dapan[vtphaydapancuoi+1:]
            #Cắt phần đã làm, set phần còn lại
            #print(macauhoi1)
            #print(pahv1)
            #print(dapan1)
            #----------
            sdvkt1 = fullcauhoi[0]['madvkt']
            sdvkt_ten1 = fullcauhoi[0]['partques']
            stieude1 = fullcauhoi[0]['titles']
            dstieude_ten = models.Monhoc.objects.filter(mamon = xmamon).values('matd1', 'tentd1', 'matd2', 'tentd2', 'matd3', 'tentd3', 'matd4', 'tentd4', 'matd5', 'tentd5')
            for it in range(5):
                if stieude1 == dstieude_ten[0][f'matd{it+1}']:
                    stieude_ten = dstieude_ten[0][f'tentd{it+1}']
                    break
                else:
                    stieude_ten = ""
            #scauhoichung1 = ""
            #Nếu các dvkt, tiêu đề trùng thì cho bằng ""
            if sdvkt1 != sdvkt:
                sdvkt = sdvkt1
                sdvkt_tendung = sdvkt_ten1
            else:
                sdvkt_tendung = ""
            if stieude1 != stieude:
                stieude = stieude1
                stieude_tendung = stieude_ten
            else:
                stieude_tendung = ""
            #--------------
            if fullcauhoi[0]['classifyques'] == "NH":
                #Câu hỏi nhóm thì questions là câu hỏi chung
                scauhoichung = fullcauhoi[0]['questions']
                option_cauhoi = fullcauhoi[0]['answer']
                ipic = fullcauhoi[0]['picques']
                iaud = fullcauhoi[0]['audques']
                iclip = fullcauhoi[0]['clipques']
                #ssdau = kq_truoc[j].answer#option_cauhoi
                vtrissdau = option_cauhoi.find("<#GP#>")#vtcachcacoption
                while vtrissdau == 0:
                    option_cauhoi = option_cauhoi[6:]
                    vtrissdau = option_cauhoi.find("<#GP#>")# Lấy vị trí của hết câu hỏi, đáp án con 1
                solgcaucon = option_cauhoi.count("<#GP#>") + 1# đếm số câu hỏi, đáp án con
                jj = 0
                # max hoasn vij ddax copy
                for n1 in range(solgcaucon):
                    if vtrissdau >= 0:
                        sscau = option_cauhoi[:vtrissdau]
                    else:
                        sscau = option_cauhoi
                    option_cauhoi = option_cauhoi[vtrissdau + 6:]
                    vtrissdau = option_cauhoi.find("<#GP#>")#để ý câu con cuối cùng không có dấu này
                    vtchung = sscau.find("[<")
                    if vtchung >= 0 and vtchung < 2:
                        vtchungcuoi = sscau.find(">]")
                        scauhoi = sscau[vtchung + 2 : vtchungcuoi]
                        sscau = sscau[vtchungcuoi + 2:]
                    else:
                        scauhoi = ""
                    #Lấy các options
                    option = {}
                    # Đếm số các phương án trả lời
                    solgpa = sscau.count("<##>") + 1
                    optiontrc = []# Giữ giá trị đáp án chưa hoán vị
                    for x in range(solgpa):
                        vtcachcacluachon = sscau.find("<##>")
                        if vtcachcacluachon > 0 :
                            ndtach = sscau[:vtcachcacluachon]#lấy option 1
                            sscau = sscau[vtcachcacluachon + 4:]#opton còn lại
                        else:
                            ndtach = sscau
                        #vtcach = ndtach.find("<##>")#vị trí phân cách các phần options của câu hỏi con
                        optiontrc.append(ndtach)
                    #lấy phương án hoán vị
                    vtpa = pahv1.find(">")
                    if vtpa >= 0 :
                        spahv = pahv1[1:vtpa]#lấy được dãy phương án hoán vị
                        sotthv = []
                        for m in range(len(spahv)):
                            sotthv.append(int(spahv[m]))
                        for n in range(len(optiontrc)):
                            option.update({nhan_dap_an[n] : optiontrc[sotthv[n]]})
                    #lấy đáp án
                    sdapan = dapan1[jj+1:jj+2]
                    #print(dapan1)
                    jj = jj + 2
                    #Nạp vào QUIZ DATA
                    assscauhoi = "<p>" + sdvkt_tendung + "</p><p>" + stieude_tendung + "</p><p>" + scauhoichung + "</p>"
                    scauhoichung = ""
                    sdvkt_tendung = ""
                    stieude_tendung = ""
                    pic = ipic
                    ipic = ""
                    aud = iaud
                    iaud = ""
                    clip = iclip
                    iclip = ""
                    sttcauhoi += 1
                    QUIZ_DATA.append({'tt':sttcauhoi, 'assquestion':assscauhoi, 'pic':pic, 'aud':aud, 'clip':clip, 'question': scauhoi, 'options': option, 'correct_answer': sdapan})
                    #print(option)
            #--------------
            elif fullcauhoi[0]['classifyques'] == "LC":
                option = {}
                #scauhoichung = ""
                scauhoi = fullcauhoi[0]['questions']
                sscau = fullcauhoi[0]['answer']
                ipic = fullcauhoi[0]['picques']
                iaud = fullcauhoi[0]['audques']
                iclip = fullcauhoi[0]['clipques']
                # Đếm số các phương án trả lời
                solgpa = sscau.count("<##>") + 1
                optiontrc = []# Giữ giá trị đáp án chưa hoán vị
                for x in range(solgpa):
                    vtcachcacluachon = sscau.find("<##>")
                    if vtcachcacluachon > 0 :
                        ndtach = sscau[:vtcachcacluachon]#lấy option 1
                        sscau = sscau[vtcachcacluachon + 4:]#opton còn lại
                    else:
                        ndtach = sscau
                    optiontrc.append(ndtach)
                #lấy phương án hoán vị
                spahv = pahv1#lấy được dãy phương án hoán vị
                sotthv = []
                for m in range(len(spahv)):
                    sotthv.append(int(spahv[m]))
                #print(sotthv)
                #print(optiontrc)
                for n in range(len(optiontrc)):
                    option.update({nhan_dap_an[n] : optiontrc[sotthv[n]]})
                #lấy đáp án
                #print(option)
                sdapan = dapan1
                #Nạp vào QUIZ DATA
                assscauhoi = "<p>" + sdvkt_tendung + "</p><p>" + stieude_tendung + "</p>"
                sdvkt_tendung = ""
                stieude_tendung = ""
                pic = ipic
                ipic = ""
                aud = iaud
                iaud = ""
                clip = iclip
                iclip = ""
                sttcauhoi += 1
                QUIZ_DATA.append({'tt':sttcauhoi, 'assquestion':assscauhoi, 'pic':pic, 'aud':aud, 'clip':clip, 'question': scauhoi, 'options': option, 'correct_answer': sdapan})
                #print(QUIZ_DATA)
        #assscauhoi = sdvkt + stieude + scauhoichung
    return QUIZ_DATA, smahvi

def khoi_tao_dl():
    QUIZ_DATA = [
        {
            'id': 1,
            'question': 'Thành phần nào trong Django chịu trách nhiệm xử lý logic và tương tác dữ liệu?',
            'options': {'A': 'Template', 'B': 'Model', 'C': 'View', 'D': 'URLconf'},
            'correct_answer': 'B'
        },
        {
            'id': 2,
            'question': 'Django tuân theo kiến trúc thiết kế nào?',
            'options': {'A': 'MVC', 'B': 'MVT', 'C': 'MVVM', 'D': 'MVP'},
            'correct_answer': 'B'
        },
        {
            'id': 3,
            'question': 'Lệnh nào dùng để chạy server phát triển của Django?',
            'options': {'A': 'python manage.py runserver', 'B': 'python manage.py startapp', 'C': 'django-admin run', 'D': 'python run.py'},
            'correct_answer': 'A'
        }
    ]
    return QUIZ_DATA

def admin_dashboard(request):
    # Nhập, sửa tên đăng nhập sinh viên
    #danhsachuser = models.Usercs.objects.all().values()
    #solguser = danhsachuser.count()
    #----------------
    context ={}
    """
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'btnSubmit':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            xuser = request.POST.get('hname')
            xuser = xuser.strip()
            xps = request.POST.get('hps')
            if xuser != "":
                # Kiểm tra có trong Usercs
                print("có rồi")
                print(xuser)
                # chuyển hướng đến một URL mới:
                return HttpResponseRedirect('../admin_dashboard.html')
            else:
                print("không có")
                print(xuser)

    else:
        context = {
            'title': 'Trang đăng nhập',
            'year' : datetime.now().year,
        }
    """
    context = {
        'title': 'Quản lý người dùng, hệ thống',
        'year' : datetime.now().year,
        'ten': ""
    }  
    return render(request, 'app/admin_dashboard.html', context)

def admin_users(request):
    # Nhập, sửa tên đăng nhập sinh viên
    danhsachuser = models.Usercs.objects.all().values()
    solguser = danhsachuser.count()
    #print(danhsachuser[0].get('id'))
    #----------------
    context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context['user_form'] = forms.UserFormrep(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['user_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                matrc = context['user_form'].cleaned_data['htendn']
                matrc = matrc.strip()
                co = False
                # kiểm tra mã có trong danh sách hay chưa
                if solguser > 0:
                    if danhsachuser[0].get('tendn') == matrc:
                        co = True
                        vit = 0
                    elif danhsachuser[len(danhsachuser)-1].get('tendn') == matrc:
                        co = True
                        vit = len(danhsachuser)-1
                    else:
                        for row in range(len(danhsachuser)):
                            if matrc == danhsachuser[row].get('tendn'):
                                co = True
                                vit = row
                                break
                    if not co:
                        utendn = context['user_form'].cleaned_data['htendn']
                        utendaydu = context['user_form'].cleaned_data['htendaydu']
                        umatkhau = context['user_form'].cleaned_data['hmatkhau']
                        uemailcn = context['user_form'].cleaned_data['hemailcn']
                        #role = 0: admin, role = 1: teacher, role = 2: sinh viên
                        if context['user_form'].cleaned_data['hrole'] == "admin":
                            urole = 0
                        elif context['user_form'].cleaned_data['hrole'] == "teacher":
                            urole = 1
                        else:
                            urole = 2
                        models.Usercs.objects.create(id = solguser + 1, tendn = utendn, tendaydu = utendaydu, matkhau = umatkhau, emailcn = uemailcn, role = urole)
                        #models.Nganhhoc.objects.filter(manganh = matrc).update(manganh = mas)
                        #x = models.Nganhhoc.objects.all()[vit]
                        #models.Nganhhoc.objects.all()[vit].manganh = mas
                        #models.Nganhhoc.objects.all()[vit].save()
                        #Dùng phương thức save sẽ thêm vào ở cuối csdl, nên dùng update()
                else:
                    utendn = context['user_form'].cleaned_data['htendn']
                    utendaydu = context['user_form'].cleaned_data['htendaydu']
                    umatkhau = context['user_form'].cleaned_data['hmatkhau']
                    uemailcn = context['user_form'].cleaned_data['hemailcn']
                    if context['user_form'].cleaned_data['hrole'] == "admin":
                        urole = 0
                    elif context['user_form'].cleaned_data['hrole'] == "teacher":
                        urole = 1
                    else:
                        urole = 2
                    models.Usercs.objects.create(id = solguser + 1, tendn = utendn, tendaydu = utendaydu, matkhau = umatkhau, emailcn = uemailcn, role = urole)
                # chuyển hướng đến một URL mới:
                return HttpResponseRedirect('/admin_users/')
        elif pressed_button == 'repair':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context['user_form'] = forms.UserFormrep(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['user_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                matrc = context['user_form'].cleaned_data['htendn']
                matrc = matrc.strip()
                co = False
                # kiểm tra mã có trong danh sách hay chưa
                if solguser > 0:
                    if danhsachuser[0].get('tendn') == matrc:
                        co = True
                        vit = 0
                    elif danhsachuser[len(danhsachuser)-1].get('tendn') == matrc:
                        co = True
                        vit = len(danhsachuser)-1
                    else:
                        for row in range(len(danhsachuser)):
                            if matrc == danhsachuser[row].get('tendn'):
                                co = True
                                vit = row
                                break
                    if co:
                        utendn = context['user_form'].cleaned_data['htendn']
                        utendaydu = context['user_form'].cleaned_data['htendaydu']
                        umatkhau = context['user_form'].cleaned_data['hmatkhau']
                        uemailcn = context['user_form'].cleaned_data['hemailcn']
                        if context['user_form'].cleaned_data['hrole'] == "admin":
                            urole = 0
                        elif context['user_form'].cleaned_data['hrole'] == "teacher":
                            urole = 1
                        else:
                            urole = 2
                        models.Usercs.objects.filter(tendn = matrc).update(tendn = utendn, tendaydu = utendaydu, matkhau = umatkhau, emailcn = uemailcn, role = urole)
                else:
                    utendn = context['user_form'].cleaned_data['htendn']
                    utendaydu = context['user_form'].cleaned_data['htendaydu']
                    umatkhau = context['user_form'].cleaned_data['hmatkhau']
                    uemailcn = context['user_form'].cleaned_data['hemailcn']
                    if context['user_form'].cleaned_data['hrole'] == "admin":
                        urole = 0
                    elif context['user_form'].cleaned_data['hrole'] == "teacher":
                        urole = 1
                    else:
                        urole = 2
                    models.Usercs.objects.create(id = solguser + 1, tendn = utendn, tendaydu = utendaydu, matkhau = umatkhau, emailcn = uemailcn, role = urole)
                # chuyển hướng đến một URL mới:
                return HttpResponseRedirect('/admin_users/')
        elif pressed_button == 'delete':
            # Xử lý logic xóa dữ liệu ở đây
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context['user_form'] = forms.UserFormrep(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['user_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                matrc = context['user_form'].cleaned_data['htendn']
                matrc = matrc.strip()
                co = False
                # kiểm tra mã có trong danh sách hay chưa
                if solguser > 0:
                    if danhsachuser[0].get('tendn') == matrc:
                        co = True
                        vit = 0
                    elif danhsachuser[len(danhsachuser)-1].get('tendn') == matrc:
                        co = True
                        vit = len(danhsachuser)-1
                    else:
                        for row in range(len(danhsachuser)):
                            if matrc == danhsachuser[row].get('tendn'):
                                co = True
                                vit = row
                                break
                    if co:
                        models.Usercs.objects.filter(tendn = matrc).delete()

            return HttpResponseRedirect('/admin_users/')
        elif pressed_button == 'cancel':
            # Xử lý logic khi hủy bỏ
            return HttpResponseRedirect('/about/')

    else:
        context = {
            'title': 'Trang nhập, sửa tên người dùng',
            'year' : datetime.now().year,
            'user_form': forms.UserFormrep(),
        }
        
    return render(request, 'app/admin_users.html', context)

def admin_student(request):
    # Nhập, sửa tên đăng nhập sinh viên
    solguser = models.Usercs.objects.count()
    #solguser = danhsachuser.count()
    #print(danhsachuser[0].get('id'))
    #print(danhsachuser[0].get('tendn'))
    #----------------
    context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context['user_form'] = forms.UserFormSV(request.POST, request.FILES)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['user_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                myfile = request.FILES['xls_file']
                # Đọc file Excel bằng Pandas
                df = pd.read_excel(myfile)
                #lấy 1 mã sv ra kiểm tra đã có trong file chưa
                co = False
                data_list = df.to_dict(orient="records") # Chuyển cả dataframe thành danh sách dict
                #print(data_list[0].get('MSSV'))
                kutendn = data_list[0].get('MSSV')
                if models.Sinhviencs.objects.filter(masv = kutendn).exists():
                    co = True
                #Nếu có thì bỏ qua
                if not co:
                    # Duyệt qua từng dòng trong DataFrame và lưu vào Database
                    for index, row in df.iterrows():
                        # Thêm vào ds sinhviencs
                        utendn = row['MSSV']
                        uhosv = row['Ho']
                        utensv = row['Ten']
                        umalop = row['fk_Ma_lop']
                        models.Sinhviencs.objects.create(masv = utendn, hosv = uhosv, tensv = utensv, fkmalop = umalop)
                        #data_list = df.to_dict(orient="records") # Chuyển cả dataframe thành danh sách dict
                        # Tạo danh sách các đối tượng model
                        #dssv = [models.Sinhviencs(**row) for row in data_list]
                        #models.Sinhviencs.objects.bulk_create(dssv)
                        #---------thêm vào đăng nhập user
                        solguser += 1
                        #utendn = row['MSSV']
                        utendaydu = row['Ho'].strip() + " " + row['Ten']
                        umatkhau = row['MK_SV']
                        urole = 2
                        #role = 0: admin, role = 1: teacher, role = 2: sinh viên
                        models.Usercs.objects.create(id = solguser, tendn = utendn, tendaydu = utendaydu, matkhau = umatkhau, role = urole)

    else:
        context = {
            'title': 'Trang nhập danh sách sinh viên',
            'year' : datetime.now().year,
            'user_form': forms.UserFormSV(),
        }
        
    return render(request, 'app/admin_student.html', context)

def admin_sys(request):
    # Cấu hình hệ thống
    
    #solguser = 0
    #----------------
    context = {
            'title': 'Trang nhập danh sách sinh viên',
            'year' : datetime.now().year,
            'user_form': forms.UserFormSV(),
        }
        
    return render(request, 'app/admin_sys.html', context)

def teacher_dashboard(request):
    # Cấu hình hệ thống
    
    solguser = 0
    #----------------
    #context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    context = {
        'title': 'Thiết lập đề thi, ngày thi',
        'year' : datetime.now().year,
        'ten': "",
    }
        
    return render(request, 'app/teacher_dashboard.html', context)

def teacher_lophp(request):
    # Tạo và nhập danh sách sv từ file excel và lớp học phần
    assert isinstance(request, HttpRequest)
    vitri = models.Monhoc.objects.count()
    ds = models.Monhoc.objects.all().values()
    dsmh = []
    for i in range(vitri):
        newdist = {'mamon' : ds[i]['mamon'], 'tenmon' : ds[i]['tenmon'], 'nganh' : ds[i]['fk_manganh'], 'trdo' : ds[i]['trinhdo'], 'tgian' : ds[i]['sogio']}
        dsmh.append(newdist)
    #----------------
    vitri = models.Lophocphan.objects.count()
    ds = models.Lophocphan.objects.all().values()
    dslophp = []
    for i in range(vitri):
        newdist = {'malophp' : ds[i]['malophp'], 'tenlophp' : ds[i]['tenlophp'], 'fkmanh' : ds[i]['fkmanh']}
        dslophp.append(newdist)
    #----------------
    context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save1':
            # tạo một instance của biểu mẫu và đưa tên môn học trong kỳ:
            context['u_form1'] = forms.NhapmonForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['u_form1'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                mahp = context['u_form1'].cleaned_data['hmamh']
                utenlophp = context['u_form1'].cleaned_data['htenlophp']
                ufkmanh = context['u_form1'].cleaned_data['hmanh']
                umota = context['u_form1'].cleaned_data['hmota']
                mahp = mahp.strip()
                co = False
                # kiểm tra mã có trong danh sách hay chưa, đề phòng trường hợp người dùng nnhập sai mã
                for row in range(len(dsmh)):
                    if mahp == dsmh[row]['mamon']:
                        co = True
                        break
                if co:
                    stt = 1
                    umalophp = mahp + f"{stt:03d}"
                    #kiểm tra có tồn tại mã lớp hp chưa
                    while models.Lophocphan.objects.filter(malophp = umalophp).exists():
                        stt +=1
                        umalophp = mahp + f"{stt:03d}"
                    utenlophp = utenlophp.strip()
                    ufkmanh = ufkmanh.strip()
                    #print(umalophp)
                    #print(utenlophp)
                    models.Lophocphan.objects.create(malophp = umalophp, tenlophp = utenlophp, mota = umota, fkmanh = ufkmanh, fkmamon = mahp)
                
                return HttpResponseRedirect('/teacher_lophp/')
        elif pressed_button == 'repair1':
            # tạo một instance của biểu mẫu và đưa tên môn học trong kỳ:
            context['u_form3'] = forms.SualhpForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['u_form3'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                umalhp = context['u_form3'].cleaned_data['hmalophp']
                utenlophp = context['u_form3'].cleaned_data['htenlophp']
                ufkmanh = context['u_form3'].cleaned_data['hmanh']
                umota = context['u_form3'].cleaned_data['hmota']
                umalhp = umalhp.strip()
                co = False
                # kiểm tra mã có trong danh sách hay chưa, đề phòng trường hợp người dùng nhập sai mã
                for row in range(len(dslophp)):
                    if umalhp == dslophp[row]['malophp']:
                        co = True
                        break
                if co:
                    models.Lophocphan.objects.filter(malophp = umalhp).update(tenlophp = utenlophp, fkmanh = ufkmanh, mota = umota)
                
                return HttpResponseRedirect('/teacher_lophp/')

        elif pressed_button == 'save2':
            # tạo một instance của biểu mẫu và nạp dữ liệu từ yêu cầu vào đó:
            context['u_form2'] = forms.VaolopHPForm(request.POST, request.FILES)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['u_form2'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                myfile = request.FILES['xls_file']
                umalhp = context['u_form2'].cleaned_data['hmalhp']
                umalhp = umalhp.strip()
                umanh = context['u_form2'].cleaned_data['hmanh']
                umanh = umanh.strip()
                # Đọc file Excel bằng Pandas
                df = pd.read_excel(myfile)
                #lấy 1 mã sv ra kiểm tra đã có trong file chưa
                co = False
                data_list = df.to_dict(orient="records") # Chuyển cả dataframe thành danh sách dict
                #print(data_list[0].get('MSSV'))
                kutendn = data_list[0].get('MSSV')
                if models.Dangkylophp.objects.filter(fkmasv = kutendn, fkmalophp = umalhp).exists():
                    co = True
                #Nếu có thì bỏ qua
                if not co:
                    # Duyệt qua từng dòng trong DataFrame và lưu vào Database
                    so = models.Dangkylophp.objects.count() + 1
                    for index, row in df.iterrows():
                        # Thêm vào ds sinhviencs
                        ufkmasv = row['MSSV']
                        ufkmalophp = umalhp
                        #fmamon = row['Ten']
                        #print(so + index)
                        #print(ufkmasv)
                        #print(ufkmalophp)
                        models.Dangkylophp.objects.create(tt = so + index, fkmasv = ufkmasv, fkmalophp = ufkmalophp, trangthai = 0)
                return HttpResponseRedirect('/teacher_lophp/')
    else:
        context = {
            'title': 'Lớp học phần',
            'year' : datetime.now().year,
            'data': dsmh,
            'datal': dslophp,
            'u_form1': forms.NhapmonForm(),
            'u_form2': forms.VaolopHPForm(),
            'u_form3': forms.SualhpForm(),
        }
        
    return render(request, 'app/teacher_lophp.html', context)

def teacher_lichthi(request):
    assert isinstance(request, HttpRequest)
    # Bổ sung lịch thi cho các lớp học phần, 2026-06-06 07:30:00
    if 'hnamhochh' in request.session:
        unamhochh = request.session.get('hnamhochh')
    else:
        unamhochh = ""
    #----------------
    if unamhochh != "":
        ds = models.Lophocphan.objects.filter(fkmanh = unamhochh).values('malophp', 'tenlophp', 'fkmamon', 'thoigianthi', 'ngaythi')
        tongso = len(ds)
        dsloptrongnam = []
        for i in range(tongso):
            if ds[i]['thoigianthi']:
                tg_string = ds[i]['thoigianthi']
            else:
                tg_string = ""
            date_string = ds[i]['ngaythi']
            #date_string = "2026-09-19T13:30"
            naive_dt = ""
            if date_string != "":
                try:
                    naive_dt = datetime.strptime(date_string, "%Y-%m-%dT%H:%M")
                except Exception as e:
                    pass
            newdist = {'tt' : i + 1, 'malophp' : ds[i]['malophp'], 'tenlophp' : ds[i]['tenlophp'], 'fkmamon' : ds[i]['fkmamon'], 'thoigianthi' : tg_string, 'ngaythi' : naive_dt}
            dsloptrongnam.append(newdist)
    #----------------
    context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save1':
            # tạo một instance của biểu mẫu và đưa tên môn học trong kỳ:
            # Duyệt qua các dữ liệu gửi lên theo ID của từng hàng
            vitri = 0
            for key, value in request.POST.items():
                if key.startswith('tglbai_'):
                    task_id = key.split('_')[1]
                    usogio = value if value else None
                    # Tìm trong dsloptrongnam với tt = task_id để lấy mã lớp hp
                if key.startswith('date_'):
                    task_id = key.split('_')[1]
                    utgian = value if value else None
                    if usogio and utgian:
                        # Tìm trong dsloptrongnam với tt = task_id để lấy mã lớp hp
                        umalophp1 = [item["malophp"] for item in dsloptrongnam if item["tt"] == vitri/2]
                        umalophp = umalophp1[0]
                        models.Lophocphan.objects.filter(malophp = umalophp).update(thoigianthi = int(usogio), ngaythi = utgian)
                vitri += 1
             
            return HttpResponseRedirect('/teacher_lichthi/')

    else:
        context = {
            'title': 'Thiết lập lịch thi',
            'year' : datetime.now().year,
            'dslop': dsloptrongnam,
        }
        
    return render(request, 'app/teacher_lichthi.html', context)

def tea_nghg(request):
    assert isinstance(request, HttpRequest)
    # Hiện các câu hỏi, bổ sung file hình ảnh, âm thanh, video cho câu hỏi ngan hàng đề thi
    #makythi = "2026-2027001"
    #Lấy mã môn học
    if 'hmonthi' in request.session:
        umamon = request.session.get('hmonthi')
    else:
        umamon = ""
    #Hiện tất cả câu hỏi theo thứ tự
    if umamon != "":
        #Lấy tên môn học đưa vào context để hiển thị
        utenmon = models.Monhoc.objects.filter(mamon = umamon).values('tenmon')
        utenmon = utenmon[0]['tenmon']
        #----------------
        ds = models.Subjects001.objects.filter(mamon = umamon).values()
        #Lọc ra đưa vào danh sách chuẩn bị hiển thị lên bảng
        tongso = len(ds)
        ds1 = []
        for i in range(tongso):
            newdist = {'mach' : ds[i]['macauhoi'], 'dvkt' : ds[i]['madvkt'], 'tieude' : ds[i]['titles'], 'cauhoi' : ds[i]['questions'], 'tam' : ds[i]['answer'], 'dapan' : ds[i]['correctans'], 'picch' : ds[i]['picques'], 'audch' : ds[i]['audques'], 'clipch' : ds[i]['clipques'], 'permch' : ds[i]['permuteques']}
            ds1.append(newdist)
        # Sort by value macauhoi
        ds1.sort(key=lambda item: item['mach'])
        #print(ds1[100]['tam'])
        #Chuyển sang danh sách hiển thị lên bảng
        tongso = len(ds1)
        dsquestions = []
        for i in range(tongso):
            #newdist = {'tt' : i + 1, 'dvkt' : ds1[i]['dvkt'], 'tieude' : ds1[i]['tieude'], 'cauhoi' : ds1[i]['tam'], 'dapan' : ds1[i]['dapan'], 'picch' : ds1[i]['picch'], 'audch' : ds1[i]['audch'], 'clipch' : ds1[i]['clipch'], 'permch' : ds1[i]['permch']}
            newdist = {'tt' : i + 1, 'dvkt' : ds1[i]['dvkt'], 'tieude' : ds1[i]['tieude'], 'cauhoi' : ds1[i]['cauhoi'], 'dapan' : ds1[i]['dapan'], 'picch' : ds1[i]['picch'], 'audch' : ds1[i]['audch'], 'clipch' : ds1[i]['clipch'], 'permch' : ds1[i]['permch']}
            dsquestions.append(newdist)
    #Form bổ sung, chỉnh sửa: gồm chỉnh sửa file âm thanh, file hình ảnh, nội dung câu hỏi chung, ...
    #Muốn chỉnh sửa thì nhập số thứ tự trên bảng câu hỏi ở trên và nhập thông tin sửa chữa.
    #Các thông tin không thay đổi sẽ không được chỉnh sửa.
    context ={}
    message = ""
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save1':
            # tạo một instance của biểu mẫu và cập nhật file hình, âm thanh hoặc xóa tên nếu kg có file
            context['g_form'] = forms.FilePicForm(request.POST, request.FILES)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['g_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                #Lấy số tt câu hỏi
                uttch1 = (context['g_form'].cleaned_data['httch'])
                try:
                    uttch = int(uttch1)
                except Exception as e:
                    uttch = -1
                #Kiểm tra file có được chọn hay không
                if 'hpic' in request.FILES:
                    # Có file được gửi lên
                    upic = request.FILES['hpic']
                    # Tạo đường dẫn lưu file trên server
                    upload_dir_abs = os.path.join(settings.MEDIA_ROOT, 'images')
                    # Kiểm tra và tạo thư mục nếu chưa có
                    os.makedirs(upload_dir_abs, exist_ok=True)
                    # Đường dẫn vật lý đầy đủ của file
                    file_path = os.path.join(upload_dir_abs, upic.name)
                    # Ghi file theo từng chunk để tối ưu bộ nhớ
                    if os.path.exists(file_path):
                        message = f"Lỗi: File '{upic.name}' đã tồn tại trên máy chủ!"
                    else:
                        with open(file_path, 'wb+') as destination:
                            for chunk in upic.chunks():
                                destination.write(chunk)
                    upic1 = upic.name
                else:
                    # Không có file nào được gửi lên
                    upic1 = ""
                #Đến file video
                if 'hclip' in request.FILES:
                    uclip = request.FILES['hclip']
                    # Tạo đường dẫn lưu file
                    upload_dir = os.path.join(settings.MEDIA_ROOT, 'videos')
                    os.makedirs(upload_dir, exist_ok=True)
                    file_path = os.path.join(upload_dir, uclip.name)
                    # Ghi file theo từng chunk để tối ưu bộ nhớ
                    if os.path.exists(file_path):
                        message = f"Lỗi: File '{uclip.name}' đã tồn tại trên máy chủ!"
                    else:
                        with open(file_path, 'wb+') as destination:
                            for chunk in uclip.chunks():
                                destination.write(chunk)
                    uclip1 = uclip.name
                else:
                    # Không có file nào được gửi lên
                    uclip1 = ""
                #Đến file âm thanh
                if 'haud' in request.FILES:
                    uaud = request.FILES['haud']
                    # Tạo đường dẫn lưu file
                    upload_dir = os.path.join(settings.MEDIA_ROOT, 'audios')
                    os.makedirs(upload_dir, exist_ok=True)
                    file_path = os.path.join(upload_dir, uaud.name)
                    # Ghi file theo từng chunk để tối ưu bộ nhớ
                    if os.path.exists(file_path):
                        message = f"Lỗi: File '{uaud.name}' đã tồn tại trên máy chủ!"
                    else:
                        with open(file_path, 'wb+') as destination:
                            for chunk in uaud.chunks():
                                destination.write(chunk)
                    uaud1 = uaud.name
                else:
                    # Không có file nào được gửi lên
                    uaud1 = ""
                #Lưu tất cả vào csdl
                if (uttch > 0) and (uttch <= tongso):
                    models.Subjects001.objects.filter(macauhoi = ds1[uttch-1]['mach']).update(picques = upic1, audques = uaud1, clipques = uclip1)
                
            return HttpResponseRedirect('/tea_nghg/')
        elif pressed_button == 'save2':
            # Cập nhật lại nội dung câu hỏi và đáp án
            context['h_form'] = forms.CauhoiFormrep(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['h_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                #Lấy số tt câu hỏi
                uttch1 = (context['h_form'].cleaned_data['httch'])
                try:
                    uttch = int(uttch1)
                except Exception as e:
                    uttch = -1
                uques = (context['h_form'].cleaned_data['hndch'])
                uopt1 = (context['h_form'].cleaned_data['hopt1'])
                uopt2 = (context['h_form'].cleaned_data['hopt2'])
                uopt3 = (context['h_form'].cleaned_data['hopt3'])
                uopt4 = (context['h_form'].cleaned_data['hopt4'])
                uopt5 = (context['h_form'].cleaned_data['hopt5'])
                uans = (context['h_form'].cleaned_data['hans'])
                if uopt1 != "":
                    uopt = uopt1
                    if uopt2 != "":
                        uopt = uopt + "<##>" + uopt2
                        if uopt3 != "":
                            uopt = uopt + "<##>" + uopt3
                            if uopt4 != "":
                                uopt = uopt + "<##>" + uopt4
                                if uopt5 != "":
                                    uopt = uopt + "<##>" + uopt5
                #print(uques)
                #print(uopt)
                #print(uans)
                #Lưu tất cả vào csdl
                if (uttch > 0) and (uttch <= tongso):
                    if (uques != "") and (uopt != "") and (uans != ""):
                        models.Subjects001.objects.filter(macauhoi = ds1[uttch-1]['mach']).update(questions = uques, answer = uopt, correctans = uans)
                
            return HttpResponseRedirect('/tea_nghg/')
    else:
        context = {
            'title': "Chỉnh, sửa ngân hàng câu hỏi",
            'mon': utenmon,
            'year' : datetime.now().year,
            'dscauhoi': dsquestions,
            'msg': message,
            'g_form': forms.FilePicForm(),
            'h_form': forms.CauhoiFormrep(),
        }
        
    return render(request, 'app/tea_nghg.html', context)

def doc_so_thap_phan(so):
    chu_so = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
    
    # Chuyển số thành chuỗi để xử lý phần nguyên và phần thập phân
    s_so = str(so)
    if '.' in s_so:
        phan_nguyen_str, phan_thap_phan_str = s_so.split('.')
    else:
        phan_nguyen_str, phan_thap_phan_str = s_so, ""
    
    # Đọc phần nguyên (đơn giản hóa đọc từng chữ số hoặc mở rộng theo ý muốn)
    ket_qua_nguyen = " ".join([chu_so[int(c)] for c in phan_nguyen_str])
    
    # Đọc phần thập phân nếu có
    if phan_thap_phan_str:
        ket_qua_thap_phan = " ".join([chu_so[int(c)] for c in phan_thap_phan_str])
        return f"{ket_qua_nguyen} phẩy {ket_qua_thap_phan}"
    
    return ket_qua_nguyen

def tea_kqt(request):
    # Tạo và nhập danh sách sv từ file excel và lớp học phần
    assert isinstance(request, HttpRequest)
    #Lấy năm học hiện hành
    xds = models.CurrentSetting.objects.filter(tt = 1).values('fkmanh')
    unhhh = xds[0]['fkmanh']
    #if 'hnamhochh' in request.session:
    #    unhhh = request.session.get('hnamhochh')
    #else:
    #    unhhh = ""
    if unhhh != "":
        ds = models.Lophocphan.objects.filter(fkmanh = unhhh).values()
        dslophp = []
        for i in range(len(ds)):
            newlist = {'malophp' : ds[i]['malophp'], 'tenlophp' : ds[i]['tenlophp'], 'fkmamon' : ds[i]['fkmamon'], 'ngaythi' : ds[i]['ngaythi']}
            dslophp.append(newlist)
    #----------------
    context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save1':
            # tạo một instance của biểu mẫu và đưa tên môn học trong kỳ:
            context['u_form'] = forms.XuatlhpForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['u_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                umalhpcb = context['u_form'].cleaned_data['hfkmalophp']
                if umalhpcb != None:
                    umalhp = umalhpcb.malophp
                    #Xuất ra excel
                    # 1. Truy vấn dữ liệu từ database
                    dskq = models.Ketquathi.objects.filter(fkmalophp = umalhp).values('id', 'fkmasv', 'fkmahoanvi', 'mark', 'ngaythi', 'fkmakt')
                    # 2. Chuyển Queryset thành Pandas DataFrame
                    df = pd.DataFrame(list(dskq))

                    if df.empty:
                        return HttpResponse("Không có dữ liệu thỏa mãn lớp học phần.", content_type="text/plain")
    
                    # Đổi tên cột cho đẹp (tùy chọn)
                    df.rename(columns={
                        'id': 'ID',
                        'fkmasv': 'Mã sinh viên',
                        'fkmahoanvi': 'Mã đề thi',
                        'mark': "Điểm",
                        'ngaythi': 'Ngày thi',
                        'fkmakt': "Kỳ thi"
                    }, inplace=True)

                    # 2. Tạo một đối tượng BytesIO để lưu file Excel trong bộ nhớ
                    with io.BytesIO() as buffer:
                        # Xuất DataFrame ra file Excel bằng engine 'openpyxl'
                        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                            df.to_excel(writer, sheet_name='Danhsach', index=False)
        
                        buffer.seek(0)
                        # 3. Tạo HttpResponse với Content-Type chuẩn để tải file Excel
                        response = HttpResponse(
                            buffer.getvalue(),
                            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                        # Đặt tên file khi tải xuống máy người dùng
                        response['Content-Disposition'] = 'attachment; filename="exketquathi.xlsx"'
                
                    return response
                else:
                    return HttpResponseRedirect('/tea_kqt/')

        elif pressed_button == 'print1':
            context['u_form'] = forms.XuatlhpForm(request.POST)
            # kiểm tra tính hợp lệ của biểu mẫu:
            if context['u_form'].is_valid():
                # xử lý dữ liệu trong form.cleaned_data theo yêu cầu
                umalhpcb = context['u_form'].cleaned_data['hfkmalophp']
                # tạo một instance của biểu mẫu và đưa tên môn học trong kỳ:
                if umalhpcb != None:
                    print(umalhpcb)
                    umalhp = umalhpcb.malophp
                    utenlhp = umalhpcb.tenlophp
                    # 1. Truy vấn dữ liệu từ database
                    ds = models.Ketquathi.objects.filter(fkmalophp = umalhp).values('id', 'fkmasv', 'fkmahoanvi', 'mark', 'ngaythi', 'fkmakt')
                    # 2. Chuyển vào ds in ra web
                    ukythi = ds[0]['fkmakt']
                    ungthi = ds[0]['ngaythi']
                    dskqin = []
                    for i in range(len(ds)):
                        newlist = {'tt': i + 1, 'fkmasv' : ds[i]['fkmasv'], 'fkmahoanvi' : ds[i]['fkmahoanvi'], 'mark' : ds[i]['mark'], 'dchu': doc_so_thap_phan(ds[i]['mark'])}
                        dskqin.append(newlist)
                    context = {
                        'kythi': ukythi,
                        'namhoc': unhhh,
                        'tenmon': utenlhp,
                        'ngthi': ungthi,
                        'data': dskqin,
                        }
                    return render(request, 'app/tea_printkqt.html', context)
                else:
                    return HttpResponse("Chưa chọn lớp học phần.", content_type="text/plain")
                    #return HttpResponseRedirect('/tea_kqt/')
    else:
        context = {
            'title': 'Xuất kết quả thi',
            'year' : datetime.now().year,
            'nh': unhhh,
            'datal': dslophp,
            'u_form': forms.XuatlhpForm(),
        }
        
    return render(request, 'app/tea_kqt.html', context)

def student_dashboard(request):
    # Cấu hình hệ thống
    
    #solguser = 0
    #----------------
    #context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    context = {
        'title': 'Bài thi trắc nghiệm',
        'year' : datetime.now().year,
        'ten': "",
    }
        
    return render(request, 'app/student_dashboard.html', context)

def student_subj(request):
    # Cấu hình hệ thống
    xds = models.CurrentSetting.objects.filter(tt = 1).values('fkmanh')
    hocky = xds[0]['fkmanh']
    #if 'hnamhochh' in request.session:
    #    hocky = request.session.get('hnamhochh')
    #else:
    #    hocky = ""
    if 'husername' in request.session:
        uusername = request.session.get('husername')
    else:
        uusername = ""
    # xoas lops hoc phan lwu trong session
    if 'hsubject' in request.session:
        del request.session['hsubject']
    #-----------------
    if uusername != "":
        ds = models.Dangkylophp.objects.filter(fkmasv = uusername).values('fkmalophp', 'trangthai')
        #ds = models.Dangkylophp.objects.all().values()
        #lấy ds các lớp đã đăng ký toàn bộ
        tongso = len(ds)
        dsloptrongnam = []
        #tongso = 0
        for i in range(tongso):
            #lọc lấy các lớp hp trong năm, trong kỳ
            dsloc = models.Lophocphan.objects.filter(malophp = ds[i]['fkmalophp']).values('malophp', 'tenlophp', 'mota', 'fkmanh', 'fkmamon', 'thoigianthi', 'ngaythi')
            tongsotrong = len(dsloc)
            for j in range(tongsotrong):
                if dsloc[j]['fkmanh'] == hocky:
                    dsloptrongnam.append(dsloc[j])
        #đưa vào ds combo các môn để hiện lên web
        solguser = len(dsloptrongnam)
        categories = []
        for x in range(solguser):
            categories.append({'tt':x, 'name' : dsloptrongnam[x]['tenlophp']})
        #----------------
        #print(categories)
    context ={}
    # nếu là yêu cầu POST, ta cần xử lý dữ liệu biểu mẫu
    if request.method == 'POST':
        # Kiểm tra người dùng nhấn nút nào
        pressed_button = request.POST.get('action')
        if pressed_button == 'save1':
            # tạo một instance của biểu mẫu và đưa tên môn học trong kỳ:
            # Duyệt qua các dữ liệu gửi lên theo ID của từng hàng
            bb = int(request.POST.get('myCombo'))
            if bb != None :
                for item in ds:
                    if (item['fkmalophp'] == dsloptrongnam[bb]['malophp']):
                        if (item['trangthai'] != 5):
                            request.session['hsubject'] = dsloptrongnam[bb]['malophp']
                            break
                        else:
                            messages.warning(request, 'Bạn không đủ điều kiện dự thi hoặc đã quá số lần thi!')
                return HttpResponseRedirect('/student_exam/')

    else:
        context = {
            'title': 'Chọn môn thi',
            'year' : datetime.now().year,
            'dslop': dsloptrongnam,
            'categories': categories,
        }
        
    return render(request, 'app/student_subj.html', context)

def student_exam(request):
    #Lấy mã lớp học phần, suy ra mã môn học; lấy mã đề chuẩn, mã đề hoán vị để xuất khởi tạo
    if 'hsubject' in request.session:
        umalophp = request.session.get('hsubject')
    else:
        umalophp = ""
    if 'husername' in request.session:
        umasv = request.session.get('husername')
    else:
        umasv = ""
    if umasv != "":
        dssv = models.Sinhviencs.objects.filter(masv = umasv).values()
        tensv = dssv[0]['hosv'] + " " + dssv[0]['tensv']
    else:
        tensv = ""
    #Khoi tao bien context ve rong
    questions = [
        {
            'tt': 1,
            'assquestion':"",
            'pic': "",
            'aud':"",
            'clip':"",
            'question': "Không thể xuất dữ liệu bài thi do Bạn chưa chọn môn thi!",
            'options': {'A': 'Tiếp tục', 'B': 'Quay lại', 'C': 'Không chọn'},
            'correct_answer': 'B'
        }
    ]
    exam_duration = 60
    #Bat dau lay du lieu
    if umalophp != "":
        ufkmamon1 = models.Lophocphan.objects.filter(malophp = umalophp).values('fkmamon', 'tenlophp', 'thoigianthi')
        ufkmamon = ufkmamon1[0]['fkmamon']
        utenlop_result = ufkmamon1[0]['tenlophp']
        if ufkmamon1[0]['thoigianthi'] != None:
            exam_duration = ufkmamon1[0]['thoigianthi'] * 60
        else:
            exam_duration = 60
        #(thời gian làm bài tính bằng phút, chuyển đổi sang giây để tính giờ)
        umadechuankynay1 = models.Cacdethikynay.objects.filter(fmamon = ufkmamon).values('madechuan', 'fmakt')
        #print(ufkmamon)
        umadechuankynay = umadechuankynay1[0]['madechuan']
        ufkmakt = umadechuankynay1[0]['fmakt']
        if 'husername' in request.session:
            uusername = request.session.get('husername')
            #del request.session['husername']
        else:
            uusername = ""
        #request.session.flush()
        # Lấy toàn bộ danh sách câu hỏi (hoặc random tùy bạn)
        #Lấy mã đề chuẩn rồi gọi hàm khoi_tao_dl_subj(mdechuan) để ra câu hỏi và đáp án in ra màn hình
        questions, umahvi = khoi_tao_dl_subj(umadechuankynay)
        # Thời gian làm bài tính bằng giây (Ví dụ: 15 phút = 900 giây)
    if request.method == 'POST':
        score = 0
        total = len(questions)
        results = []
        correct_ans_num = 0
        udapanchon = ""
        stt = models.Ketquathi.objects.count()
        for q in questions:
            # Lấy đáp án thí sinh chọn từ form gửi lên
            user_choice = request.POST.get(f'question_{q["tt"]}')
            is_correct = (user_choice == q['correct_answer'])
            if is_correct:
                correct_ans_num += 1
            results.append({
                'question': q['question'],
                'user_choice': user_choice,
                'correct_ans': q['correct_answer'],
                'is_correct': is_correct,
                'correct_ans_num': correct_ans_num,
                'total': total,
                'is_passed': True,
            })
            if user_choice == None :
                udapanchon = udapanchon + ","
            else:
                udapanchon = udapanchon + "," + user_choice
        #lưu vào bảng kết quả thi
        stt += 1
        umark = round(correct_ans_num * 10/total,1)
        now = datetime.now()
        ungaythi = now.strftime("%Y-%m-%dT%H:%M")
        #print(ungaythi)
        #ss = {'id' : stt, 'fkmalophp' : umalophp, 'fkmasv' : uusername, 'fkmamon' : ufkmamon, 'fkmadechuan' : umadechuankynay, 'fkmahoanvi' : umahvi, 'dapanchon' : udapanchon, 'mark' : umark, 'ngaythi' : ungaythi, 'fkmakt' : ufkmakt}
        #print(ss)
        models.Ketquathi.objects.create(id = stt, fkmalophp = umalophp, fkmasv = uusername, fkmamon = ufkmamon, fkmadechuan = umadechuankynay, fkmahoanvi = umahvi, dapanchon = udapanchon, mark = umark, ngaythi = ungaythi, fkmakt = ufkmakt)
        # Lưu vào lớp hoc phần trang thái đã thi lần n
        ds = models.Dangkylophp.objects.filter(fkmasv = uusername, fkmalophp = umalophp).values('trangthai')
        lanthi = int(ds[0]['trangthai'])
        if lanthi < 4:
            lanthi += 1
        #print(lanthi)
        models.Dangkylophp.objects.filter(fkmasv = uusername, fkmalophp = umalophp).update(trangthai=lanthi)
        #models.Ketquathi.objects.all().delete()
        if umark >= 5:
            is_passed = True
        else:
            is_passed = False
        context = {
            'title': utenlop_result,
            'completed_at': now,
            'score': umark,
            'correct_ans_num': correct_ans_num,
            'total': total,
            'results': results,
            'is_passed': is_passed,
        }
        return render(request, 'app/exam_result.html', context)

    context = {
        'tenmon': utenlop_result,
        'hoten': tensv,
        'questions': questions,
        'exam_duration': exam_duration,
        'is_submitted': False,
    }
    return render(request, 'app/student_exam.html', context)
