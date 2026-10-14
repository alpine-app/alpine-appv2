"""
Definition of forms.
"""

from django import forms
from .models import Nganhhoc, Monhoc, Dvktm, Namhoc, Kythi, Lophocphan
from django.utils.translation import gettext_lazy as _

class MonhocForm(forms.Form):
    hmamon = forms.CharField(label="Mã môn học:", max_length=20)
    htenmon = forms.CharField(label="Tên môn học:", max_length=255, required=False)
    hfk_manganh = forms.ModelChoiceField(queryset=Nganhhoc.objects.all(), label='Danh mục ngành', empty_label='-- Chọn ngành --')
    htrinhdo = forms.CharField(label="Trình độ:", max_length=50, required=False)
    hsogio = forms.IntegerField(label="Số giờ:", required=False)
    hmadvkt1 = forms.CharField(label="Mã đơn vị kiến thức 1:", max_length=20, required=False)
    htendvkt1 = forms.CharField(label="Tên đơn vị kiến thức/ Chương/ Phần 1:", max_length=500, required=False)
    hmadvkt2 = forms.CharField(label="Mã đơn vị kiến thức 2:", max_length=20, required=False)
    htendvkt2 = forms.CharField(label="Tên đơn vị kiến thức/ Chương/ Phần 2:", max_length=500, required=False)
    hmadvkt3 = forms.CharField(label="Mã đơn vị kiến thức 3:", max_length=20, required=False)
    htendvkt3 = forms.CharField(label="Tên đơn vị kiến thức/ Chương/ Phần 3:", max_length=500, required=False)
    hmadvkt4 = forms.CharField(label="Mã đơn vị kiến thức 4:", max_length=20, required=False)
    htendvkt4 = forms.CharField(label="Tên đơn vị kiến thức/ Chương/ Phần 4:", max_length=500, required=False)
    hmadvkt5 = forms.CharField(label="Mã đơn vị kiến thức 5:", max_length=20, required=False)
    htendvkt5 = forms.CharField(label="Tên đơn vị kiến thức/ Chương/ Phần 5:", max_length=500, required=False)
    hmatd1 = forms.CharField(label="Mã tiêu đề 1:", max_length=20, required=False)
    htentd1 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 300,'class': 'form-control','placeholder': 'Nhập nội dung...'}), label="Nội dung tiêu đề 1:", required=False)
    hmatd2 = forms.CharField(label="Mã tiêu đề 2:", max_length=20, required=False)
    htentd2 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 300,'class': 'form-control'}), label="Nội dung tiêu đề 2:", required=False)
    hmatd3 = forms.CharField(label="Mã tiêu đề 3:", max_length=20, required=False)
    htentd3 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 300,'class': 'form-control'}), label="Nội dung tiêu đề 3:", required=False)
    hmatd4 = forms.CharField(label="Mã tiêu đề 4:", max_length=20, required=False)
    htentd4 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 300,'class': 'form-control'}), label="Nội dung tiêu đề 4:", required=False)
    hmatd5 = forms.CharField(label="Mã tiêu đề 5:", max_length=20, required=False)
    htentd5 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 300,'class': 'form-control'}), label="Nội dung tiêu đề 5:", required=False)

class MonhocFormrep(forms.Form):
    hmamoned1 = forms.CharField(label="Mã môn học cần sửa:", max_length=20)
    hmamoned = forms.CharField(label="Sửa mã môn học:", max_length=20)
    htenmoned = forms.CharField(label="Sửa tên môn học:", max_length=255, required=False)
    hfk_manganhed = forms.ModelChoiceField(queryset=Nganhhoc.objects.all(), label='Danh mục ngành', empty_label='-- Chọn ngành --')
    htrinhdoed = forms.CharField(label="Trình độ:", max_length=50, required=False)
    hsogioed = forms.IntegerField(label="Số giờ:", required=False)
    hmadvkt1ed = forms.CharField(label="Mã đơn vị kiến thức 1:", max_length=20, required=False)
    htendvkt1ed = forms.CharField(label="Tên đơn vị kiến thức/ Chương/ Phần 1:", max_length=500, required=False)
    hmadvkt2ed = forms.CharField(label="Mã đơn vị kiến thức 2:", max_length=20, required=False)
    htendvkt2ed = forms.CharField(label="Tên đơn vị kiến thức/ Chương/ Phần 2:", max_length=500, required=False)
    hmadvkt3ed = forms.CharField(label="Mã đơn vị kiến thức 3:", max_length=20, required=False)
    htendvkt3ed = forms.CharField(label="Tên đơn vị kiến thức/ Chương/ Phần 3:", max_length=500, required=False)
    hmadvkt4ed = forms.CharField(label="Mã đơn vị kiến thức 4:", max_length=20, required=False)
    htendvkt4ed = forms.CharField(label="Tên đơn vị kiến thức/ Chương/ Phần 4:", max_length=500, required=False)
    hmadvkt5ed = forms.CharField(label="Mã đơn vị kiến thức 5:", max_length=20, required=False)
    htendvkt5ed = forms.CharField(label="Tên đơn vị kiến thức/ Chương/ Phần 5:", max_length=500, required=False)
    hmatd1ed = forms.CharField(label="Mã tiêu đề 1:", max_length=20, required=False)
    htentd1ed = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 300,'class': 'form-control','placeholder': 'Nhập nội dung...'}), label="Nội dung tiêu đề 1:", required=False)
    hmatd2ed = forms.CharField(label="Mã tiêu đề 2:", max_length=20, required=False)
    htentd2ed = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 300,'class': 'form-control'}), label="Nội dung tiêu đề 2:", required=False)
    hmatd3ed = forms.CharField(label="Mã tiêu đề 3:", max_length=20, required=False)
    htentd3ed = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 300,'class': 'form-control'}), label="Nội dung tiêu đề 3:", required=False)
    hmatd4ed = forms.CharField(label="Mã tiêu đề 4:", max_length=20, required=False)
    htentd4ed = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 300,'class': 'form-control'}), label="Nội dung tiêu đề 4:", required=False)
    hmatd5ed = forms.CharField(label="Mã tiêu đề 5:", max_length=20, required=False)
    htentd5ed = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 300,'class': 'form-control'}), label="Nội dung tiêu đề 5:", required=False)

class NganhhocForm(forms.Form):
    hmanganh = forms.CharField(label="Mã ngành học:", max_length=10)
    htennganh = forms.CharField(label="Tên ngành học:", max_length=100, required=False)
    htrinhdo = forms.CharField(label="Bậc trình độ:", max_length=50, required=False)
    hkhoa = forms.CharField(label="Khoa quản lý:", max_length=100, required=False)

class NganhhocFormrep(forms.Form):
    hmanganhed1 = forms.CharField(label="Mã ngành học cần sửa:", max_length=10)
    hmanganhed = forms.CharField(label="Nhập lại mã ngành:", max_length=10, required=False)
    htennganhed = forms.CharField(label="Nhập lại tên ngành:", max_length=100, required=False)
    htrinhdoed = forms.CharField(label="Nhập lại bậc trình độ:", max_length=50, required=False)
    hkhoaed = forms.CharField(label="Nhập lại Khoa quản lý:", max_length=100, required=False)

class MatranchondeForm(forms.Form):
    # Bảng gồm Mã, tên, tổng số câu, số câu từng mức độ hiểu biết (RD, DE, TB, KH, RK)
    xmadvkt = forms.CharField(label="Nhập mã đơn vị kiến thức", max_length=10)
    xtendvkt = forms.CharField(label="Nhập tên đơn vị kiến thức", max_length=100)

class CauhoiForm(forms.Form):
    httch = forms.CharField(label="Thứ tự câu hỏi:", max_length=10, required=False)
    hndch = forms.CharField(widget=forms.Textarea(attrs={'rows': 6,'cols': 1000,'class': 'form-control','placeholder': 'Nhập nội dung...'}), label="Nội dung câu hỏi:", required=False)
    hpic = forms.FileField(label="Tên file hình ảnh:", required=False)
    haud = forms.FileField(label="Chọn file âm thanh", required=False)
    hclip = forms.FileField(label="Tên file video:", required=False)
    hopt1 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 1000,'class': 'form-control','placeholder': 'Lựa chọn 1...'}), label="A.", required=False)
    hpic1 = forms.FileField(label="Tên file hình ảnh:", required=False)
    hopt2 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 1000,'class': 'form-control','placeholder': 'Lựa chọn 2...'}), label="B.", required=False)
    hpic2 = forms.FileField(label="Tên file hình ảnh:", required=False)
    hopt3 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 1000,'class': 'form-control','placeholder': 'Lựa chọn 3...'}), label="C.", required=False)
    hpic3 = forms.FileField(label="Tên file hình ảnh:", required=False)
    hopt4 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 1000,'class': 'form-control','placeholder': 'Lựa chọn 4...'}), label="D.", required=False)
    hpic4 = forms.FileField(label="Tên file hình ảnh:", required=False)
    hopt5 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 1000,'class': 'form-control','placeholder': 'Lựa chọn 5...'}), label="E.", required=False)
    hpic5 = forms.FileField(label="Tên file hình ảnh:", required=False)
    hans = forms.CharField(label="Đáp án đúng:", max_length=10, required=False)
    #roll_number = forms.IntegerField(help_text = "Enter 6 digit roll number")
    #password = forms.CharField(widget = forms.PasswordInput(), required=False)

class FileCHForm(forms.Form):
    hfile = forms.FileField(label="Chọn file word nhập câu hỏi", required=False)

class MakeExamForm(forms.Form):
    hmakt = forms.CharField(label="Mã kỳ thi:", max_length=20, required=False)
    htenkt = forms.CharField(label="Tên kỳ thi:", max_length=255, required=False)
    hfk_manh = forms.ModelChoiceField(queryset=Namhoc.objects.all(), label='Năm học', empty_label='-- Chọn năm học --', required=False)
    hmotakt = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 300,'class': 'form-control','placeholder': 'Kỳ thi này là...'}), label="Mô tả kỳ thi:", required=False)

class GetMatrandeForm(forms.Form):
    hfk_mamtr = forms.CharField(label="Nhập mã ma trận đề trong số các mã trên:", max_length=20, required=False)

class XuatKythiForm(forms.Form):
    hfk_makt = forms.ModelChoiceField(queryset=Kythi.objects.all(), label='Kỳ thi', empty_label='-- Chọn kỳ thi --', required=False)
    #hfk_manh = forms.ModelChoiceField(queryset=Namhoc.objects.all(), label='Năm học', empty_label='-- Chọn năm học --', required=False)

class XuatCombomonForm(forms.Form):
    hmavten = forms.ModelChoiceField(queryset=Monhoc.objects.all(), label='Danh mục môn học', required=False, empty_label='-- Chọn môn học --')

class XuatdvktmForm(forms.Form):
    mavtenm = forms.ModelChoiceField(queryset=Dvktm.objects.all(), label='Danh mục đơn vị kiến thức', required=False, empty_label='-- Chọn đơn vị kiến thức --')

class UserFormrep(forms.Form):
    htendn = forms.CharField(label="Tên đăng nhập:", max_length=100, required=False)
    htendaydu = forms.CharField(label="Tên đầy đủ:", max_length=300, required=False)
    hmatkhau = forms.CharField(label="Mật khẩu:", max_length=255, required=False)
    hemailcn = forms.CharField(label="Email:", max_length=200, required=False)
    hrole = forms.ChoiceField(choices=[("admin", "Quản trị"), ("teacher", "Giáo viên"), ("student", "Sinh viên"),], label="Chọn phân quyền", widget=forms.Select,)

class UserFormSV(forms.Form):
    xls_file = forms.FileField(label='Chọn file Excel')

class NhapmonForm(forms.Form):
    hmamh = forms.CharField(label="Mã môn học:", max_length=20, required=False)
    htenlophp = forms.CharField(label="Tên lớp học phần:", max_length=255, required=False)
    hmanh = forms.CharField(label="Mã năm học:", max_length=10, required=False)
    hmota = forms.CharField(label="Mô tả:", max_length=500, required=False)

class SualhpForm(forms.Form):
    hmalophp = forms.CharField(label="Mã lớp học phần:", max_length=30, required=False)
    htenlophp = forms.CharField(label="Tên lớp học phần:", max_length=255, required=False)
    hmanh = forms.CharField(label="Mã năm học:", max_length=10, required=False)
    hmota = forms.CharField(label="Mô tả:", max_length=500, required=False)

class VaolopHPForm(forms.Form):
    xls_file = forms.FileField(label='Chọn file Excel')
    hmalhp = forms.CharField(label="Mã lớp học phần:", max_length=30, required=False)
    hmanh = forms.CharField(label="Mã học kỳ:", max_length=10, required=False)

class FilePicForm(forms.Form):
    httch = forms.CharField(label="Nhập số thứ tự câu hỏi cần chỉnh sửa:", max_length=10, required=False)
    hpic = forms.FileField(label="Tên file hình ảnh:", required=False)
    haud = forms.FileField(label="Chọn file âm thanh", required=False)
    hclip = forms.FileField(label="Tên file video:", required=False)

class CauhoiFormrep(forms.Form):
    httch = forms.CharField(label="Thứ tự câu hỏi:", max_length=10, required=False)
    hndch = forms.CharField(widget=forms.Textarea(attrs={'rows': 6,'cols': 1000,'class': 'form-control','placeholder': 'Nhập nội dung...'}), label="Nội dung câu hỏi:", required=False)
    hopt1 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 1000,'class': 'form-control','placeholder': 'Lựa chọn 1...'}), label="A.", required=False)
    hopt2 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 1000,'class': 'form-control','placeholder': 'Lựa chọn 2...'}), label="B.", required=False)
    hopt3 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 1000,'class': 'form-control','placeholder': 'Lựa chọn 3...'}), label="C.", required=False)
    hopt4 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 1000,'class': 'form-control','placeholder': 'Lựa chọn 4...'}), label="D.", required=False)
    hopt5 = forms.CharField(widget=forms.Textarea(attrs={'rows': 2,'cols': 1000,'class': 'form-control','placeholder': 'Lựa chọn 5...'}), label="E.", required=False)
    hans = forms.CharField(label="Đáp án đúng:", max_length=10, required=False)
    
class NamhocForm(forms.Form):
    hmanh = forms.CharField(label="Mã năm học:", max_length=10, required=False)
    hmahk1 = forms.CharField(label="Năm bắt đầu:", max_length=10, required=False)
    hmahk2 = forms.CharField(label="Năm kết thúc:", max_length=10, required=False)

class NamhochhForm(forms.Form):
    hmanh = forms.CharField(label="Mã năm học:", max_length=10, required=False)

class XuatlhpForm(forms.Form):
    hfkmalophp = forms.ModelChoiceField(queryset=Lophocphan.objects.all(), label='Lớp học phần', empty_label='-- Chọn lớp --', required=False)

