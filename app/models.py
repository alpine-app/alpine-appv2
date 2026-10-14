"""
Definition of models.
"""

from django.db import models

class AuthGroup(models.Model):
    name = models.CharField(unique=True, max_length=150)

    class Meta:
        managed = False
        db_table = 'auth_group'


class AuthPermission(models.Model):
    name = models.CharField(max_length=255)
    content_type = models.ForeignKey('DjangoContentType', models.DO_NOTHING)
    codename = models.CharField(max_length=100)

    class Meta:
        managed = False
        db_table = 'auth_permission'
        unique_together = (('content_type', 'codename'),)


class AuthGroupPermissions(models.Model):
    id = models.BigAutoField(primary_key=True)
    group = models.ForeignKey(AuthGroup, models.DO_NOTHING)
    permission = models.ForeignKey('AuthPermission', models.DO_NOTHING)

    class Meta:
        managed = False
        db_table = 'auth_group_permissions'
        unique_together = (('group', 'permission'),)


class AuthUser(models.Model):
    password = models.CharField(max_length=128)
    last_login = models.DateTimeField(blank=True, null=True)
    is_superuser = models.BooleanField()
    username = models.CharField(unique=True, max_length=150)
    first_name = models.CharField(max_length=150)
    last_name = models.CharField(max_length=150)
    email = models.CharField(max_length=254)
    is_staff = models.BooleanField()
    is_active = models.BooleanField()
    date_joined = models.DateTimeField()

    class Meta:
        managed = False
        db_table = 'auth_user'


class AuthUserGroups(models.Model):
    id = models.BigAutoField(primary_key=True)
    user = models.ForeignKey(AuthUser, models.DO_NOTHING)
    group = models.ForeignKey(AuthGroup, models.DO_NOTHING)

    class Meta:
        managed = False
        db_table = 'auth_user_groups'
        unique_together = (('user', 'group'),)


class AuthUserUserPermissions(models.Model):
    id = models.BigAutoField(primary_key=True)
    user = models.ForeignKey(AuthUser, models.DO_NOTHING)
    permission = models.ForeignKey(AuthPermission, models.DO_NOTHING)

    class Meta:
        managed = False
        db_table = 'auth_user_user_permissions'
        unique_together = (('user', 'permission'),)


class DjangoAdminLog(models.Model):
    action_time = models.DateTimeField()
    object_id = models.TextField(blank=True, null=True)
    object_repr = models.CharField(max_length=200)
    action_flag = models.SmallIntegerField()
    change_message = models.TextField()
    content_type = models.ForeignKey('DjangoContentType', models.DO_NOTHING, blank=True, null=True)
    user = models.ForeignKey(AuthUser, models.DO_NOTHING)

    class Meta:
        managed = False
        db_table = 'django_admin_log'


class DjangoContentType(models.Model):
    app_label = models.CharField(max_length=100)
    model = models.CharField(max_length=100)

    class Meta:
        managed = False
        db_table = 'django_content_type'
        unique_together = (('app_label', 'model'),)


class DjangoMigrations(models.Model):
    id = models.BigAutoField(primary_key=True)
    app = models.CharField(max_length=255)
    name = models.CharField(max_length=255)
    applied = models.DateTimeField()

    class Meta:
        managed = False
        db_table = 'django_migrations'


class DjangoSession(models.Model):
    session_key = models.CharField(primary_key=True, max_length=40)
    session_data = models.TextField()
    expire_date = models.DateTimeField()

    class Meta:
        managed = False
        db_table = 'django_session'

class Namhoc(models.Model):
    manh = models.CharField(primary_key=True, max_length=10)
    namhk1 = models.CharField(max_length=10, blank=True, null=True)
    namhk2 = models.CharField(max_length=10, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'namhoc'

class Kythi(models.Model):
    makt = models.CharField(primary_key=True, max_length=20)
    tenkythi = models.CharField(max_length=200, blank=True, null=True)
    motakt = models.TextField(blank=True, null=True)
    fkmanh = models.CharField(max_length=10, blank=True, null=True)
    ngaybatdau = models.CharField(max_length=25, blank=True, null=True)
    ngayketthuc = models.CharField(max_length=25, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'kythi'

class Cacdethi(models.Model):
    madechuan = models.CharField(primary_key=True, max_length=20)
    fmamon = models.CharField(max_length=20, blank=True, null=True)
    fmakt = models.CharField(max_length=20, blank=True, null=True)
    noidungchuan = models.TextField(blank=True, null=True)
    dapanchuan = models.CharField(max_length=300, blank=True, null=True)
    mahoanvi1 = models.CharField(max_length=10, blank=True, null=True)
    noidunghv1 = models.TextField(blank=True, null=True)
    dapanhv1 = models.CharField(max_length=300, blank=True, null=True)
    pahv1 = models.TextField(blank=True, null=True)
    mahoanvi2 = models.CharField(max_length=10, blank=True, null=True)
    noidunghv2 = models.TextField(blank=True, null=True)
    dapanhv2 = models.CharField(max_length=300, blank=True, null=True)
    pahv2 = models.TextField(blank=True, null=True)
    mahoanvi3 = models.CharField(max_length=10, blank=True, null=True)
    noidunghv3 = models.TextField(blank=True, null=True)
    dapanhv3 = models.CharField(max_length=300, blank=True, null=True)
    pahv3 = models.TextField(blank=True, null=True)
    mahoanvi4 = models.CharField(max_length=10, blank=True, null=True)
    noidunghv4 = models.TextField(blank=True, null=True)
    dapanhv4 = models.CharField(max_length=300, blank=True, null=True)
    pahv4 = models.TextField(blank=True, null=True)
    mahoanvi5 = models.CharField(max_length=10, blank=True, null=True)
    noidunghv5 = models.TextField(blank=True, null=True)
    dapanhv5 = models.CharField(max_length=300, blank=True, null=True)
    pahv5 = models.TextField(blank=True, null=True)
    mahoanvi6 = models.CharField(max_length=10, blank=True, null=True)
    noidunghv6 = models.TextField(blank=True, null=True)
    dapanhv6 = models.CharField(max_length=300, blank=True, null=True)
    pahv6 = models.TextField(blank=True, null=True)
    mahoanvi7 = models.CharField(max_length=10, blank=True, null=True)
    noidunghv7 = models.TextField(blank=True, null=True)
    dapanhv7 = models.CharField(max_length=300, blank=True, null=True)
    pahv7 = models.TextField(blank=True, null=True)
    mahoanvi8 = models.CharField(max_length=10, blank=True, null=True)
    noidunghv8 = models.TextField(blank=True, null=True)
    dapanhv8 = models.CharField(max_length=300, blank=True, null=True)
    pahv8 = models.TextField(blank=True, null=True)
    mahoanvi9 = models.CharField(max_length=10, blank=True, null=True)
    noidunghv9 = models.TextField(blank=True, null=True)
    dapanhv9 = models.CharField(max_length=300, blank=True, null=True)
    pahv9 = models.TextField(blank=True, null=True)
    pahvcau1 = models.CharField(max_length=255, blank=True, null=True)
    pahvcau2 = models.CharField(max_length=255, blank=True, null=True)
    pahvcau3 = models.CharField(max_length=255, blank=True, null=True)
    pahvcau4 = models.CharField(max_length=255, blank=True, null=True)
    pahvcau5 = models.CharField(max_length=255, blank=True, null=True)
    pahvcau6 = models.CharField(max_length=255, blank=True, null=True)
    pahvcau7 = models.CharField(max_length=255, blank=True, null=True)
    pahvcau8 = models.CharField(max_length=255, blank=True, null=True)
    pahvcau9 = models.CharField(max_length=255, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'cacdethi'

class Cacdethikynay(models.Model):
    madechuan = models.CharField(primary_key=True, max_length=20)
    fmamon = models.CharField(max_length=20, blank=True, null=True)
    fmakt = models.CharField(max_length=20, blank=True, null=True)
    lanthi = models.SmallIntegerField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'cacdethikynay'


class Nganhhoc(models.Model):
    manganh = models.CharField(primary_key=True, max_length=10)
    tennganh = models.CharField(max_length=100, blank=True, null=True)
    trinhdo = models.CharField(max_length=50, blank=True, null=True)
    khoa = models.CharField(max_length=100, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'nganhhoc'


class Tabletitles(models.Model):
    mavtitles = models.TextField(primary_key=True)
    fk_mamon = models.CharField(max_length=20, blank=True, null=True)
    tenmon = models.CharField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'tabletitles'


class Matranchonde(models.Model):
    mamtr = models.CharField(primary_key=True, max_length=20)
    fk_mamon = models.CharField(max_length=20, blank=True, null=True)
    noidungmtr = models.TextField(blank=True, null=True)
    socaumtr = models.SmallIntegerField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'matranchonde'


class Monhoc(models.Model):
    mamon = models.CharField(primary_key=True, max_length=20)
    tenmon = models.CharField(blank=True, null=True)
    fk_manganh = models.CharField(max_length=20, blank=True, null=True)
    trinhdo = models.CharField(max_length=50, blank=True, null=True)
    sogio = models.SmallIntegerField(blank=True, null=True)
    madvkt1 = models.CharField(max_length=20, blank=True, null=True)
    tendvkt1 = models.TextField(blank=True, null=True)
    madvkt2 = models.CharField(max_length=20, blank=True, null=True)
    tendvkt2 = models.TextField(blank=True, null=True)
    madvkt3 = models.CharField(max_length=20, blank=True, null=True)
    tendvkt3 = models.TextField(blank=True, null=True)
    madvkt4 = models.CharField(max_length=20, blank=True, null=True)
    tendvkt4 = models.TextField(blank=True, null=True)
    madvkt5 = models.CharField(max_length=20, blank=True, null=True)
    tendvkt5 = models.TextField(blank=True, null=True)
    madvkt6 = models.CharField(max_length=20, blank=True, null=True)
    tendvkt6 = models.TextField(blank=True, null=True)
    madvkt7 = models.CharField(max_length=20, blank=True, null=True)
    tendvkt7 = models.TextField(blank=True, null=True)
    madvkt8 = models.CharField(max_length=20, blank=True, null=True)
    tendvkt8 = models.TextField(blank=True, null=True)
    madvkt9 = models.CharField(max_length=20, blank=True, null=True)
    tendvkt9 = models.TextField(blank=True, null=True)
    madvkt10 = models.CharField(max_length=20, blank=True, null=True)
    tendvkt10 = models.TextField(blank=True, null=True)
    matd1 = models.CharField(max_length=20, blank=True, null=True)
    tentd1 = models.TextField(blank=True, null=True)
    matd2 = models.CharField(max_length=20, blank=True, null=True)
    tentd2 = models.TextField(blank=True, null=True)
    matd3 = models.CharField(max_length=20, blank=True, null=True)
    tentd3 = models.TextField(blank=True, null=True)
    matd4 = models.CharField(max_length=20, blank=True, null=True)
    tentd4 = models.TextField(blank=True, null=True)
    matd5 = models.CharField(max_length=20, blank=True, null=True)
    tentd5 = models.TextField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'monhoc'


class Dvktm(models.Model):
    mavtendvkt = models.TextField(primary_key=True)
    fk_mamon = models.CharField(max_length=20, blank=True, null=True)
    tenmon = models.CharField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'dvktm'


class Sinhviencs(models.Model):
    masv = models.CharField(primary_key=True, max_length=100)
    hosv = models.CharField(max_length=300, blank=True, null=True)
    tensv = models.CharField(max_length=100, blank=True, null=True)
    fkmalop = models.CharField(max_length=30, blank=True, null=True)
    emailsv = models.CharField(max_length=100, blank=True, null=True)
    dtsv = models.CharField(max_length=20, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'sinhviencs'

class Lophocphan(models.Model):
    malophp = models.CharField(primary_key=True, max_length=30)
    tenlophp = models.CharField(max_length=255, blank=True, null=True)
    mota = models.CharField(max_length=500, blank=True, null=True)
    fkmanh = models.CharField(max_length=10, blank=True, null=True)
    fkmamon = models.CharField(max_length=20, blank=True, null=True)
    thoigianthi = models.IntegerField(blank=True, null=True)
    ngaythi = models.CharField(max_length=25, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'lophocphan'

class Subjects001(models.Model):
    macauhoi = models.CharField(primary_key=True, max_length=20)
    titles = models.CharField(max_length=20, blank=True, null=True)
    questions = models.TextField(blank=True, null=True)
    picques = models.CharField(max_length=100, blank=True, null=True)
    audques = models.CharField(max_length=100, blank=True, null=True)
    clipques = models.CharField(max_length=100, blank=True, null=True)
    permuteques = models.BooleanField(blank=True, null=True)
    answer = models.TextField(blank=True, null=True)
    picans = models.TextField(blank=True, null=True)
    audans = models.TextField(blank=True, null=True)
    clipans = models.TextField(blank=True, null=True)
    permuteans = models.CharField(max_length=250, blank=True, null=True)
    classifyques = models.CharField(max_length=20, blank=True, null=True)
    partques = models.CharField(max_length=100, blank=True, null=True)
    levelques = models.CharField(max_length=20, blank=True, null=True)
    correctans = models.CharField(max_length=50, blank=True, null=True)
    layoutans1 = models.SmallIntegerField(blank=True, null=True)
    mamon = models.CharField(max_length=20, blank=True, null=True)
    madvkt = models.CharField(max_length=10, blank=True, null=True)
    createdate = models.CharField(max_length=25, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'subjects001'


class Usercs(models.Model):
    id = models.IntegerField(primary_key=True)
    tendn = models.CharField(max_length=100, blank=True, null=True)
    tendaydu = models.CharField(max_length=300, blank=True, null=True)
    matkhau = models.CharField(max_length=255, blank=True, null=True)
    emailcn = models.CharField(max_length=200, blank=True, null=True)
    role = models.SmallIntegerField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'usercs'


class Dangkylophp(models.Model):
    tt = models.IntegerField(primary_key=True)
    fkmasv = models.CharField(max_length=100, blank=True, null=True)
    fkmalophp = models.CharField(max_length=30, blank=True, null=True)
    trangthai = models.SmallIntegerField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'dangkylophp'


class Ketquathi(models.Model):
    id = models.IntegerField(primary_key=True)
    fkmalophp = models.CharField(max_length=30, blank=True, null=True)
    fkmasv = models.CharField(max_length=100, blank=True, null=True)
    fkmamon = models.CharField(max_length=20, blank=True, null=True)
    fkmadechuan = models.CharField(max_length=20, blank=True, null=True)
    fkmahoanvi = models.CharField(max_length=10, blank=True, null=True)
    dapanchon = models.TextField(blank=True, null=True)
    mark = models.FloatField(blank=True, null=True)
    ngaythi = models.CharField(max_length=25, blank=True, null=True)
    fkmakt = models.CharField(max_length=20, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'ketquathi'

class CurrentSetting(models.Model):
    tt = models.SmallIntegerField(primary_key=True)
    fkmanh = models.CharField(max_length=10, blank=True, null=True)
    makt = models.CharField(max_length=20, blank=True, null=True)
    
    class Meta:
        managed = False
        db_table = 'currentsetting'
