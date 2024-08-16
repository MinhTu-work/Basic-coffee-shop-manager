from abc import ABC, abstractmethod
from MSSQLConnection import MSSQLConnection

#Tạo class abcNhanVien trừu tượng
class abcNhanVien(ABC):
    @abstractmethod
    def tinhluongHT(self):
        pass

#Tạo class NhanVien kế thừa class abcNhanVien trừu tượng
class NhanVien(abcNhanVien):
    def __init__(self, maNV, **kwargs):
        self.__maNV = maNV
        self.__hoTen = kwargs.get('hoTen', 'Cap nhat sau')
        self.__luongCB = kwargs.get('luongCB', 0)
        self.__luongHT = 0

    #Đóng gói
    @property
    def maNV(self):
        return self.__maNV

    @property
    def hoTen(self):
        return self.__hoTen

    @property
    def luongCB(self):
        return self.__luongCB

    @property
    def luongHT(self):
        return self.__luongHT

    @luongHT.setter
    def luongHT(self, luongHT):
        self.__luongHT = luongHT

    def tinhluongHT(self):
        pass

#Tạo class Nhân viên văn phòng (NVVP) kế thừa class NhanVien
class NVVP(NhanVien):
    def __init__(self, maNV, **kwargs):
        super().__init__(maNV, hoTen = kwargs.get('hoTen', 'Cap nhat sau'),
                         luongCB = kwargs.get('luongCB', 0))
        self.__soNG = kwargs.get('soNG', 0)
        self.congViec = kwargs.get('congViec', 'Cap nhat sau')

    def tinhLuongHT(self):
        luong = self.luongCB + self.__soNG * 150_000
        self.luongHT = luong
        return luong

#Tạo class Nhân viên bán hàng (NVBH) kế thừa class NhanVien
class NVBH(NhanVien):
    def __init__(self, maNV, **kwargs):
        super().__init__(maNV, hoTen = kwargs.get('hoTen', 'Cap nhat sau'),
                         luongCB = kwargs.get('luongCB', 0))
        self.__soSP = kwargs.get('soSP')
        self.congViec = kwargs.get('congViec', 'Cap nhat sau')

    def tinhLuongHT(self):
        luong = self.luongCB + self.__soSP * 18_000
        self.luongHT = luong
        return luong

#Tạo class CongTy, để thực hiện các method lên nhân viên
class CongTy:
    def __init__(self, maCT, tenCT):
        self.__maCT = maCT
        self.__tenCT = tenCT
        self.__ds = []

    #Đóng gói
    @property
    def ds(self):
        return self.__ds

    #Khởi tạo các nhân viên từ CSDL
    def init_ds_nv(self):
        cn = MSSQLConnection()
        cn.connect()

        rows = cn.query(
            "select NhanVien.MaNhanVien, HoTen, LuongCoBan, SoNgayLam, SoSanPham, LoaiNhanVien from NhanVien, ChamCongTongHop where NhanVien.MaNhanVien = ChamCongTongHop.MaNhanVien")
        for row in rows:
            if row[5] == 'Bán Hàng':
                nv2 = NVBH(row[0], hoTen=row[1], luongCB=row[2], soSP=row[4], congViec=row[5])
                self.__ds.extend([nv2])
            else:
                nv1 = NVVP(row[0], hoTen=row[1], luongCB=row[2], soNG=row[3], congViec=row[5])
                self.__ds.extend([nv1])

        cn.close()

    #Tính lương hàng tháng mỗi nhân viên
    def tinh_luong_ht(self):
        for nv in self.__ds:
            nv.tinhLuongHT()



