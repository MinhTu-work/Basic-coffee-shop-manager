import tkinter as tk
from tkinter import ttk
from MSSQLConnection import MSSQLConnection
from ChucNang import CongTy

#Class UI bao gồm Giao diện, Câu truy vấn SQL, gọi các chức năng từ các file .py khác
class UI(tk.Tk):

    #Các biến class để lưu dữ liệu cho các method
    rows = []
    rows1 = []
    rows2 = []
    sum2 = 0

    #Contructor
    def __init__(self):
        super().__init__()

        #Phần giao diện:
        #Tạo 5 tab: frame1, frame2, frame3, frame4, fram5
        self.title('Quản lý quán Cafe')
        self.notebook = ttk.Notebook(self, width=820, height=600)
        self.notebook.pack(padx=5, pady=5)
        frame1 = ttk.Frame(self.notebook)
        frame2 = ttk.Frame(self.notebook)
        frame3 = ttk.Frame(self.notebook)
        frame4 = ttk.Frame(self.notebook)
        frame5 = ttk.Frame(self.notebook)
        self.notebook.add(frame1, text="Nhân viên")
        self.notebook.add(frame2, text="Tính lương")
        self.notebook.add(frame3, text="Thanh toán")
        self.notebook.add(frame4, text="Hóa đơn")
        self.notebook.add(frame5, text="Khách hàng")


        #Khi chuyển tab thì sẽ kích hoạt hàm on_tab_change
        self.notebook.bind('<<NotebookTabChanged>>', self.on_tab_change)


        #1 Giao diện tab Nhân viên (frame1)
        self.ma_nv1 = tk.Label(master=frame1, text='Mã nhân viên: ')
        self.ma_nv1.place(x=10, y=350)
        self.ma_nv_input1 = tk.Entry(master=frame1)
        self.ma_nv_input1.place(x=100, y=350)

        self.ten_nv = tk.Label(master=frame1, text='Tên nhân viên: ')
        self.ten_nv.place(x=10, y=400)
        self.add_ten_nv = tk.Entry(master=frame1)
        self.add_ten_nv.place(x=100, y=400)

        self.luong_cb = tk.Label(master=frame1, text='Lương cơ bản: ')
        self.luong_cb.place(x=10, y=450)
        self.add_luong_cb = tk.Entry(master=frame1)
        self.add_luong_cb.place(x=100, y=450)

            #Tạo select option cho loại nhân viên
        self.loai_nv = tk.Label(master=frame1, text='Loại nhân viên: ')
        self.loai_nv.place(x=10, y=500)
        self.list_loai_nv = ("N'Văn Phòng'", "N'Bán Hàng'", "N'Không Loại'")
        self.option_loai_nv = tk.StringVar()
        self.option_loai_nv.set(self.list_loai_nv[2])
        self.add_loai_nv = tk.OptionMenu(frame1, self.option_loai_nv, *self.list_loai_nv)
        self.add_loai_nv.place(x=115, y=495)

        self.cong_lam = tk.Label(master=frame1, text='Số ngày làm /sản phẩm: ')
        self.cong_lam.place(x=10, y=550)
        self.add_cong_lam = tk.Entry(master=frame1, width=10)
        self.add_cong_lam.place(x=170, y=550)

        self.ma_nv = tk.Label(master=frame1, text='Mã nhân viên: ')
        self.ma_nv.place(x=310, y=350)
        self.ma_nv_input = tk.Entry(master=frame1)
        self.ma_nv_input.place(x=400, y=350)

        self.load_btn = tk.Button(master=frame1, text='Load', command=self.load)
        self.load_btn.place(x=10, y=250)

        self.add_btn = tk.Button(master=frame1, text='Add', command=self.add)
        self.add_btn.place(x=10, y=300)

        self.del_btn = tk.Button(master=frame1, text='Delete', command=self.delete1)
        self.del_btn.place(x=310, y=300)

        self.find_btn = tk.Button(master=frame1, text='Find', command=self.find)
        self.find_btn.place(x=410, y=300)

        self.upd_btn = tk.Button(master=frame1, text='Update', command=self.update1)
        self.upd_btn.place(x=110, y=300)

        #1.1 Tạo bảng để xuất dữ liệu trên giao diện tab Nhân viên (frame1)
        self.tree1 = ttk.Treeview(master=frame1)
        self.tree1["column"] = ("ma_nv", "ten_nv", "luong_cb", "loai_nv", "ngay_lam", "san_pham")
        self.tree1.column("#0", width=50)
        self.tree1.column("ma_nv", width=100)
        self.tree1.column("ten_nv", width=150)
        self.tree1.column("luong_cb", width=150)
        self.tree1.column("loai_nv", width=150)
        self.tree1.column("ngay_lam", width=100)
        self.tree1.column("san_pham", width=100)

        self.tree1.heading("#0", text="STT")
        self.tree1.heading("ma_nv", text="Mã nhân viên")
        self.tree1.heading("ten_nv", text="Tên nhân viên")
        self.tree1.heading("luong_cb", text="Lương cơ bản")
        self.tree1.heading("loai_nv", text="Loại nhân viên")
        self.tree1.heading("ngay_lam", text="Số ngày làm")
        self.tree1.heading("san_pham", text="Số sản phẩm")
        self.tree1.place(x=10, y=20)


        #2 Giao diện tab Tính lương (frame2)
        self.ma_nv2 = tk.Label(master=frame2, text='Mã nhân viên: ')
        self.ma_nv2.place(x=10, y=350)
        self.ma_nv_input2 = tk.Entry(master=frame2)
        self.ma_nv_input2.place(x=100, y=350)

        self.salary_btn = tk.Button(master=frame2, text='Salary', command=self.salary)
        self.salary_btn.place(x=10, y=250)

        self.search_btn = tk.Button(master=frame2, text='Search', command=self.search)
        self.search_btn.place(x=10, y=300)
        self.detached_items = []

        self.back_btn = tk.Button(master=frame2, text='Back', command=self.unhide)
        self.back_btn.place(x=110, y=300)

        #2.1 Tạo bảng để xuất dữ liệu trên giao diện tab Tính lương (frame2)
        self.tree2 = ttk.Treeview(master=frame2)
        self.tree2["column"] = ("ma_nv", "ten_nv", "loai_nv", "luong_thang")
        self.tree2.column("#0", width=50)
        self.tree2.column("ma_nv", width=100)
        self.tree2.column("ten_nv", width=150)
        self.tree2.column("loai_nv", width=150)
        self.tree2.column("luong_thang", width=150)

        #2.2 Tạo bảng và tên hàm sort theo cột trên giao diện tab Tính lương (frame2)
        self.image = tk.PhotoImage(file = '1.png')
        self.tree2.heading("#0", text="STT", command=lambda c="#0": self.sort_treeview(c, False))
        self.tree2.heading("ma_nv", text="Mã nhân viên", command=lambda c="ma_nv": self.sort_treeview(c, False), image = self.image)
        self.tree2.heading("ten_nv", text="Tên nhân viên", command=lambda c="ten_nv": self.sort_treeview(c, False), image = self.image)
        self.tree2.heading("loai_nv", text="Loại nhân viên", command=lambda c="loai_nv": self.sort_treeview(c, False), image = self.image)
        self.tree2.heading("luong_thang", text="Lương tháng", command=lambda c="luong_thang": self.sort_treeview(c, False), image = self.image)
        self.tree2.place(x=10, y=20)


        #3 Giao diện tab Thanh toán (frame3)
        self.ngay_lap = tk.Label(master=frame3, text='Ngày lập')
        self.ngay_lap.place(x=120, y=10)
        self.ngay_lap_input = tk.Entry(master=frame3)
        self.ngay_lap_input.place(x=120, y=30)

        self.ma_kh_tt = tk.Label(master=frame3, text='Mã khách hàng')
        self.ma_kh_tt.place(x=255, y=10)
        self.ma_kh_tt_input = tk.Entry(master=frame3)
        self.ma_kh_tt_input.place(x=255, y=30)

        self.ma_nv_tt = tk.Label(master=frame3, text='Mã nhân viên')
        self.ma_nv_tt.place(x=390, y=10)
        self.ma_nv_tt_input = tk.Entry(master=frame3)
        self.ma_nv_tt_input.place(x=390, y=30)

            #Tạo select option cho sản phẩm
        self.ten_sp_tt = tk.Label(master=frame3, text='Tên sản phẩm')
        self.ten_sp_tt.place(x=10, y=340)
        cn = MSSQLConnection()
        cn.connect()
        self.list_ten_sp_tt = cn.query("select MaSanPham, TenSanPham, GiaTien from SANPHAM")
        cn.close()
        self.list_ten_sp_only_tt = []
        for i in self.list_ten_sp_tt:
            self.list_ten_sp_only_tt.append(i[1])
        self.option_ten_sp_tt = tk.StringVar()
        self.option_ten_sp_tt.set(self.list_ten_sp_only_tt[0])
        self.add_ten_sp_tt = tk.OptionMenu(frame3, self.option_ten_sp_tt, *self.list_ten_sp_only_tt)
        self.add_ten_sp_tt.place(x=10, y=360)

        self.so_luong_tt = tk.Label(master=frame3, text='Số lượng')
        self.so_luong_tt.place(x=170, y=340)
        self.so_luong_tt_input = tk.Entry(master=frame3)
        self.so_luong_tt_input.place(x=170, y=360)

        self.add_tt_btn = tk.Button(master=frame3, text='Add', command=self.add_sp_tt)
        self.add_tt_btn.place(x=330, y=360)

        self.done_tt_btn = tk.Button(master=frame3, text='Chốt đơn', command=self.add_new_hd)
        self.done_tt_btn.place(x=430, y=360)

        #3.1 Tạo bảng 1 để xuất dữ liệu trên giao diện tab Thanh toán (frame3)
        self.tree3a = ttk.Treeview(master=frame3, height = 1)
        self.tree3a["column"] = ("tong_tien")
        self.tree3a['show'] = 'headings'
        self.tree3a.column("tong_tien", width=100)
        self.tree3a.heading("tong_tien", text="Tổng tiền")
        self.tree3a.place(x=600, y=5)

        #3.2 Tạo bảng 2 để xuất dữ liệu trên giao diện tab Thanh toán (frame3)
        self.tree3b = ttk.Treeview(master=frame3, height=1)
        self.tree3b["column"] = ("ma_hd")
        self.tree3b['show'] = 'headings'
        self.tree3b.column("ma_hd", width=100)
        self.tree3b.heading("ma_hd", text="Mã hóa đơn")
        self.tree3b.place(x=10, y=5)

        #3.3 Tạo bảng 3 để xuất dữ liệu trên giao diện tab Thanh toán (frame3)
        self.tree3 = ttk.Treeview(master=frame3)
        self.tree3["column"] = ("ma_sp", "ten_sp", "so_luong", "don_gia", "thanh_tien")
        self.tree3['show'] = 'headings'
        self.tree3.column("ma_sp", width=100)
        self.tree3.column("ten_sp", width=100)
        self.tree3.column("so_luong", width=100)
        self.tree3.column("don_gia", width=100)
        self.tree3.column("thanh_tien", width=100)

        self.tree3.heading("ma_sp", text="Mã sản phẩm")
        self.tree3.heading("ten_sp", text="Tên sản phẩm")
        self.tree3.heading("so_luong", text="Số lượng")
        self.tree3.heading("don_gia", text="Đơn giá")
        self.tree3.heading("thanh_tien", text="Thành tiền")
        self.tree3.place(x=10, y=100)


        #4 Giao diện tab Hóa đơn (frame4)
        self.load_btn = tk.Button(master=frame4, text='Load', command=self.load_hd)
        self.load_btn.place(x=10, y=250)

        self.ma_hd = tk.Label(master=frame4, text='Mã hóa đơn: ')
        self.ma_hd.place(x=10, y=550)
        self.ma_hd_input = tk.Entry(master=frame4)
        self.ma_hd_input.place(x=100, y=550)

        self.show_btn = tk.Button(master=frame4, text='Chi tiết Hóa đơn', command=self.show_hd)
        self.show_btn.place(x=250, y=550)
        self.show_btn = tk.Button(master=frame4, text='Delete', command=self.delete_hd)
        self.show_btn.place(x=400, y=550)

        #4.1 Tạo bảng 1 để xuất dữ liệu trên giao diện tab Hóa đơn (frame4)
        self.tree4 = ttk.Treeview(master=frame4)
        self.tree4["column"] = ("ma_hd", "ngay_lap", "ma_kh", "ma_nv", "tri_gia")
        self.tree4.column("#0", width=50)
        self.tree4.column("ma_hd", width=100)
        self.tree4.column("ngay_lap", width=150)
        self.tree4.column("ma_kh", width=150)
        self.tree4.column("ma_nv", width=150)
        self.tree4.column("tri_gia", width=100)

        self.tree4.heading("#0", text="STT")
        self.tree4.heading("ma_hd", text="Mã hóa đơn")
        self.tree4.heading("ngay_lap", text="Ngày lập")
        self.tree4.heading("ma_kh", text="Mã khách hàng")
        self.tree4.heading("ma_nv", text="Mã nhân viên")
        self.tree4.heading("tri_gia", text="Trị giá")
        self.tree4.place(x=10, y=20)

        #4.2 Tạo bảng 2 để xuất dữ liệu trên giao diện tab Hóa đơn (frame4)
        self.tree4a = ttk.Treeview(master=frame4)
        self.tree4a["column"] = ("ma_sp", "ten_sp", "so_luong", "gia_tien")
        self.tree4a.column("#0", width=50)
        self.tree4a.column("ma_sp", width=100)
        self.tree4a.column("ten_sp", width=150)
        self.tree4a.column("so_luong", width=150)
        self.tree4a.column("gia_tien", width=150)

        self.tree4a.heading("#0", text="STT")
        self.tree4a.heading("ma_sp", text="Mã sản phẩm")
        self.tree4a.heading("ten_sp", text="Tên sản phẩm")
        self.tree4a.heading("so_luong", text="Số lượng")
        self.tree4a.heading("gia_tien", text="Giá tiền")
        self.tree4a.place(x=10, y=300)


        #5 Giao diện tab Khách hàng (frame5)
        self.load_kh_btn = tk.Button(master=frame5, text='Load', command=self.load_kh)
        self.load_kh_btn.place(x=10, y=250)
        self.top5_kh_btn = tk.Button(master=frame5, text='Top 5 Mua hàng nhiều nhất', command=self.top5_kh)
        self.top5_kh_btn.place(x=110, y=250)

        self.makh_kh = tk.Label(master=frame5, text='Mã khách hàng: ')
        self.makh_kh.place(x=10, y=300)
        self.makh_kh_input = tk.Entry(master=frame5)
        self.makh_kh_input.place(x=100, y=300)

        self.find_kh_btn = tk.Button(master=frame5, text='Find', command=self.find_kh)
        self.find_kh_btn.place(x=240, y=300)

        self.delete_kh_btn = tk.Button(master=frame5, text='Delete', command=self.delete_kh)
        self.delete_kh_btn.place(x=340, y=300)

        self.tenkh_kh = tk.Label(master=frame5, text='Tên khách hàng: ')
        self.tenkh_kh.place(x=10, y=350)
        self.tenkh_kh_input = tk.Entry(master=frame5)
        self.tenkh_kh_input.place(x=100, y=350)

        self.sdtkh_kh = tk.Label(master=frame5, text='Số điện thoại: ')
        self.sdtkh_kh.place(x=10, y=400)
        self.sdtkh_kh_input = tk.Entry(master=frame5)
        self.sdtkh_kh_input.place(x=100, y=400)

        self.add_kh_btn = tk.Button(master=frame5, text='Add', command=self.add_kh)
        self.add_kh_btn.place(x=240, y=350)

        #5 Tạo bảng để xuất dữ liệu trên giao diện tab Khách hàng (frame5)
        self.tree5 = ttk.Treeview(master=frame5)
        self.tree5["column"] = ("ma_kh", "ten_kh", "sdt", "doanh_thu")
        self.tree5['show'] = 'headings'
        self.tree5.column("ma_kh", width=100)
        self.tree5.column("ten_kh", width=150)
        self.tree5.column("sdt", width=150)
        self.tree5.column("doanh_thu", width=150)
        self.tree5.heading("ma_kh", text="Mã khách hàng")
        self.tree5.heading("ten_kh", text="Tên khách hàng")
        self.tree5.heading("sdt", text="Số điện thoại")
        self.tree5.heading("doanh_thu", text="Doanh thu")
        self.tree5.place(x=10, y=20)

    #Phần chức năng
    #1. Chức năng refresh, tab Nhân viên: Nạp lại dữ liệu từ CSDL vào 1 mảng trong Python. Đảm bảo luôn cập nhật dữ liệu mới nhất
    def refresh(self):
        cn = MSSQLConnection()
        cn.connect()
        UI.rows = cn.query(
            "select NhanVien.MaNhanVien, HoTen, LuongCoBan, LoaiNhanVien, SoNgayLam, SoSanPham from NhanVien, ChamCongTongHop where NhanVien.MaNhanVien = ChamCongTongHop.MaNhanVien")
        cn.close()


    #1.1 Chức năng load, tab Nhân viên: Đưa dữ liệu từ 1 mảng trong Python lên giao diện
    def load(self):
        #Xóa hiển thị hiện hữu trên bảng (nếu có)
        for item in self.tree1.get_children():
            self.tree1.delete(item)

        #Nạp dữ liệu mới nhất từ CSDL
        self.refresh()

        #Tạo STT đếm tự động và đưa dữ liệu lên giao diện
        a = 0
        for i in UI.rows:
            a += 1
            self.tree1.insert(parent="", index='end', text=a, value=(i[0], i[1], i[2], i[3], i[4], i[5]))


    #1.2 Chức năng add, tab Nhân viên: Thêm 1 nhân viên mới vào CSDL, mã số nhân viên tự động tăng 1 đơn vị
    def add(self):
        #Lấy dữ liệu từ giao diện
        ten_nv = self.add_ten_nv.get()
        luong_cb = self.add_luong_cb.get()
        loai_nv = self.option_loai_nv.get()
        cong_lam = self.add_cong_lam.get()

        #Kiểm tra là nhân viên văn phòng hay bán hàng, để gán đúng công làm
        if loai_nv == self.list_loai_nv[1]:
            so_ngay_lam = str(0)
            so_sp = str(cong_lam)
        else:
            so_ngay_lam = str(cong_lam)
            so_sp = str(0)

        #Kiểm tra mã nhân viên hiện có trong CSDL, từ đó cộng thêm 1 để tạo mã cho nhân viên mới
        a = len(UI.rows) - 1
        c = int((UI.rows[a][0][2] + UI.rows[a][0][3])) + 1
        b = 'NV' + str(c)

        #Đưa câu lệnh sql xuống CSDL
        cn = MSSQLConnection()
        cn.connect()
        cn.insert("insert into NhanVien values('" + b + "'" + ", " + "N'" + ten_nv + "'" + ", " + luong_cb + ")")
        cn.insert("insert into ChamCongTongHop values('" + b + "'" + ", " + loai_nv + ", " + so_ngay_lam + ", " + so_sp + ")")
        cn.close()

        #Mục đích tự tải lại CSDL lên giao diện sau khi thêm 1 nhân viên mới
        self.load()


    #1.3 Chức năng update1, tab Nhân viên: Cập nhật thông tin của nhân viên vào CSDL, để trống là xóa dữ liệu ô thông tin đó.
    #Bắt buộc phải nhập Mã nhân viên, thông báo nếu Mã nhân viên không tồn tại
    def update1(self):
        #Lấy dữ liệu mã nhân viên từ giao diện
        get0 = self.ma_nv_input1.get()

        #Kiểm tra mã nhân viên có tồn tại không
        check = 0
        for row in UI.rows:
            if row[0] == get0:
                check = 1

        #Nếu mã nhân viên tồn tại, thực hiện đưa câu lệnh sql update xuống CSDL
        if check:
            #Lấy dữ liệu từ giao diện
            get1 = self.add_ten_nv.get()
            get2 = self.add_luong_cb.get()
            get3 = self.option_loai_nv.get()
            get4 = self.add_cong_lam.get()

            #Kiểm tra là nhân viên văn phòng hay bán hàng, để gán đúng công làm
            if get3 == self.list_loai_nv[1]:
                get5 = str(0)
                get6 = str(get4)
            else:
                get5 = str(get4)
                get6 = str(0)

            #Đưa câu lệnh sql xuống CSDL
            cn = MSSQLConnection()
            cn.connect()
            cn.update(
                "update NhanVien " + "set HoTen = N'" + get1 + "', " + "LuongCoBan = " + get2 + " where MaNhanVien = '" + get0 + "'")
            cn.update(
                "update ChamCongTongHop " + "set LoaiNhanVien = " + get3 + ", SoNgayLam = " + get5 + ", SoSanPham = " + get6 + " where MaNhanVien ='" + get0 + "'")
            cn.close()

            #Mục đích tự tải lại CSDL lên giao diện sau khi cập nhật 1 nhân viên
            self.load()

        #Nếu mã nhân viên không tồn tại, thì thông báo và không làm gì cả
        else:
            #Xóa hiển thị hiện hữu trên bảng (nếu có)
            for item in self.tree1.get_children():
                self.tree1.delete(item)

            #Thông báo mã nhân viên không tồn tại lên giao diện
            self.tree1.insert(parent="", index='end', text=1, value=("Không tồn tại", "", "", ""))


    #1.4 Chức năng delete1, tab Nhân viên: Xóa nhân viên ra khỏi CSDL bằng Mã nhân viên. Nếu không tìm thấy sẽ thông báo.
    def delete1(self):
        #Lấy dữ liệu mã nhân viên từ giao diện
        get1 = self.ma_nv_input.get()

        #Kiểm tra mã nhân viên có tồn tại không
        check = 0
        for row in UI.rows:
            if row[0] == get1:
                check = 1

        #Nếu mã nhân viên tồn tại, thì đưa câu lệnh sql xóa xuống CSDL
        if check:
            cn = MSSQLConnection()
            cn.connect()
            cn.delete("delete from ChamCongTongHop where MaNhanVien = '" + get1 + "'")
            cn.delete("delete from NhanVien where MaNhanVien = '" + get1 + "'")
            cn.close()

            #Mục đích tự tải lại CSDL lên giao diện sau khi xóa 1 nhân viên
            self.load()

        #Nếu mã nhân viên không tồn tại, thì thông báo và không làm gì cả
        else:
            #Xóa hiển thị hiện hữu trên bảng (nếu có)
            for item in self.tree1.get_children():
                self.tree1.delete(item)

            #Thông báo mã nhân viên không tồn tại lên giao diện
            self.tree1.insert(parent="", index='end', text=1, value=("Không tồn tại", "", "", ""))


    #1.5 Chức năng find, tab Nhân viên: Tìm 1 nhân viên trong CSDL theo Mã nhân viên. Nếu không tìm thấy sẽ thông báo
    def find(self):
        #Lấy dữ liệu mã nhân viên từ giao diện
        get1 = self.ma_nv_input.get()

        #Tìm mã nhân viên
        check = 0
        index = 0
        for row in UI.rows:
            #Nếu tìm thấy
            if row[0] == get1:
                check = 1
                #Lưu index dòng dữ liệu chứa mã nhân viên đó
                index = UI.rows.index(row)

        #Xóa hiển thị hiện hữu trên bảng (nếu có)
        for item in self.tree1.get_children():
            self.tree1.delete(item)

        #Nếu tìm thấy thì in dòng dữ liệu của mã nhân viên đó lên giao diện
        if check:
            self.tree1.insert(parent="", index='end', text=1,
                             value=(UI.rows[index][0], UI.rows[index][1], UI.rows[index][2], UI.rows[index][3],
                                    UI.rows[index][4], UI.rows[index][5]))

        #Nếu không tìm thấy thì thông báo
        else:
            self.tree1.insert(parent="", index='end', text=1, value=('Không tìm thấy', "", "", ""))



    #2. Chức năng tính lương, tab Tính lương, import từ ChucNang.py
    def salary(self):
        #Khởi tạo danh sách nhân viên công ty
        ct = CongTy(369, "Tesla")
        ct.init_ds_nv()

        #Tính lương
        ct.tinh_luong_ht()
        a = 0

        #Xóa hiện thị trên bảng
        for item in self.tree2.get_children():
            self.tree2.delete(item)

        #Hiện thị lên bảng
        for i in ct.ds:
            a += 1
            self.tree2.insert(parent="", index='end', text=a, value=(i.maNV, i.hoTen, i.congViec, i.luongHT))


    #2.1 Chức năng Sắp xếp dữ liệu trong cột tăng dần/ giảm dần, trong tab Tính lương
    #Tham số đầu vào là col (Cột cần sort), descending (Có xếp giảm dần không)
    def sort_treeview(self, col, descending):
        #Lưu dữ liệu trong cột vàd row vào biến data
        data = [(self.tree2.set(item, col), item) for item in self.tree2.get_children('')]

        #Hàm sort của python
        data.sort(reverse=descending)

        #Xếp lại dòng dữu liệu trên giao diện
        for index, (val, item) in enumerate(data):
            self.tree2.move(item, '', index)

        #Bấm lại lần nữa vào cột này thì sẽ sort ngược lại
        self.tree2.heading(col, command=lambda: self.sort_treeview(col, not descending))


    #2.2 Chức năng tìm mã nhân viên, tab Tính lương. Sử dụng hàm ẩn và hiện các dòng trueview trên giao diện
    def search(self):
        self.unhide()
        get2 = self.ma_nv_input2.get()
        temp = ''

        #Tìm mã nhân viên
        for item in self.tree2.get_children():
            if get2 == self.tree2.item(item)['values'][0]:
                temp = item

        #Chỉ hiện mã nhân viên tìm thấy
        if temp:
            self.hide()
            self.tree2.reattach(temp, '', 0)
        else:
            self.hide()


    #2.3 Ẩn các dòng trong bảng TrueView, tab Tính lương
    def hide(self):
        d = self.tree2.get_children()
        for i in d:
            self.tree2.detach(i)
            self.detached_items.append(i)


    #2.4 Hiện các dòng trong bảng TrueView, tab Tính lương
    def unhide(self):
        for i in reversed (self.detached_items):
            self.tree2.reattach(i, '', 0)
        self.detached_items.clear()


    #3. Chức năng thêm sản phẩm vào đơn hàng và tính tiền, tab Thanh toán
    #Khi thêm một món hàng và số lượng, tự động tính tiền từng sản phẩm và tổng tiền phải trả
    def add_sp_tt(self):
        #Lấy dữ liệu t giao diện
        get1 = self.option_ten_sp_tt.get()
        get2 = int(self.so_luong_tt_input.get())

        #Tìm đơn giá của sản phẩm để tính tiền sản phẩm đó (sum1), đồng thời tự cộng thêm vào tổng tiền phải trả (sum2)
        for i in self.list_ten_sp_tt:
            #Tìm được sản phẩm, tra ra đơn giá
            if get1 == i[1]:
                #Lấy số lượng nhân đơn giá
                sum1 = get2*i[2]
                #Tự cộng vào giá tiền tổng phải trả
                UI.sum2 = UI.sum2 + sum1
                #In tên sản phẩm, số lượng và tiền sản phẩm lên giao diện
                self.tree3.insert(parent="", index='end', value=(i[0], i[1], get2, i[2], sum1))
                #In tổng tiền phải trả lên giao diện
                    #Xóa tổng tiền cũ (nếu có)
                for item in self.tree3a.get_children():
                    self.tree3a.delete(item)
                    #In tổng tiền mới
                self.tree3a.insert(parent="", index='end', value=(UI.sum2))


    #3.1 Khi chuyển sang tab Thanh toán, thì sẽ chạy fuction này để tự tạo ra mã hóa đơn mới cho đơn hàng sắp chốt
    def on_tab_change(self, event):
        tab = event.widget.tab('current')['text']
        if tab == "Thanh toán":
            self.new_ma_hd()


    #3.2 Tự động tăng mã hóa đơn lên 1, khi có hóa đơn mới
    def new_ma_hd(self):
        #Xuất toàn bộ hóa đơn vào mảng
        cn = MSSQLConnection()
        cn.connect()
        UI.rows1 = cn.query("select * from HOADON")
        cn.close()

        #Tạo ra mã số hóa đơn mới (tăng thêm 1 từ hóa đơn gần đây nhất)
        a = len(UI.rows1) - 1
        c = int((UI.rows1[a][0][2] + UI.rows1[a][0][3])) + 1
        b = 'HD' + str(c)

        #In mã hóa đơn mới lên giao diện
        for item in self.tree3b.get_children():
            self.tree3b.delete(item)
        self.tree3b.insert(parent="", index='end', value=(b))


    #3.3 Thêm hóa đơn mới vào CSDL
    def add_new_hd(self):
        #Lấy dữ liệu từ giao diện
        item = self.tree3b.get_children()
        ma_hd = self.tree3b.item(item)['values'][0]

        ngay_lap = self.ngay_lap_input.get()
        ma_kh = self.ma_kh_tt_input.get()
        ma_nv = self.ma_nv_tt_input.get()

        item = self.tree3a.get_children()
        tri_gia = self.tree3a.item(item)['values'][0]

        #Đưa câu lệnh sql xuống CSDL
        cn = MSSQLConnection()
        cn.connect()
            #Set format ngày tháng cho CSDL
        cn.insert("SET DATEFORMAT dmy")
            #Tạo hóa đơn mới cho CSDL
        cn.insert("INSERT INTO HOADON(MaHoaDon, NgayLap, MaKhachHang, MaNhanVien, TriGia) VALUES ('" + ma_hd + "','" + ngay_lap + "','" + ma_kh + "','" + ma_nv + "'," + tri_gia + ")")
            #Tạo chi tiết hóa đơn mới cho CSDL
                #Lấy dữ liệu từ giao diện, có bao nhiêu dòng dữu liệu sản phẩm đặt mua thì tạo bấy nhiêu câu truy vấn
        for i in self.tree3.get_children():
            ma_sp = self.tree3.item(i)['values'][0]
            so_luong = str(self.tree3.item(i)['values'][2])
            cn.insert("INSERT INTO CTHD(MaHoaDon, MaSanPham, SoLuong) VALUES ('" + ma_hd + "','" + ma_sp + "'," + so_luong + ")")
        cn.close()

        #Tự tạo mã hóa đơn mới (Cộng 1) cho đơn hàng tiếp theo
        self.new_ma_hd()


    #4 Chức năng refresh_hd, tab Hóa đơn: Đưa dữ liệu vào mảng
    def refresh_hd(self):
        cn = MSSQLConnection()
        cn.connect()
        UI.rows1 = cn.query("select * from HOADON")
        cn.close()


    #4.1 Chức năng load_hd, tab Hóa đơn: Đưa dữ liệu từ mảng lên giao diện
    def load_hd(self):
        #Xóa hiển thị trên bảng (nếu có)
        for item in self.tree4.get_children():
            self.tree4.delete(item)

        #Nạp dữ liệu mới từ CSDL vào mảng
        self.refresh_hd()

        #a là cột số STT tự dộng công 1 trên trueview
        a = 0
        for i in UI.rows1:
            a += 1
            #Đưa dữ liệu lên giao diện
            self.tree4.insert(parent="", index='end', text=a, value=(i[0], i[1], i[2], i[3], i[4]))


    #4.2 Chức năng show_hd, tab Hóa đơn: Tìm Mã hóa đơn và xuất chi tiết hóa đơn đó
    def show_hd(self):
        #Lấy mã hóa đơn từ giao diện
        get = self.ma_hd_input.get()

        check = 0
        index = 0

        #Nạp dữ liệu mới từ CSDL vào mảng
        self.refresh_hd()

        #Tìm mã hóa đơn có tồn tại trong CSDL không
        for row in UI.rows1:
            #Nếu tìm thấy
            if row[0] == get:
                check = 1
                #Lưu index dòng dữ liệu của mã hóa đơn
                index = UI.rows1.index(row)

        #Xóa hiện thị trên bảng hóa đơn (nếu có)
        for item in self.tree4.get_children():
            self.tree4.delete(item)

        #Nếu tìm thấy thì đưa dữ liệu hóa đơn lên bảng
        if check:
            self.tree4.insert(parent="", index='end', text=1,
                              value=(UI.rows1[index][0], UI.rows1[index][1], UI.rows1[index][2], UI.rows1[index][3],
                                     UI.rows1[index][4]))

            #Lấy dữ liệu về chi tiết hóa đơn từ CSDL
            cn = MSSQLConnection()
            cn.connect()
            hd1 = cn.query(
                "select cthd.MaSanPham, sp.TenSanPham, SoLuong, GiaTien from CTHD join SANPHAM sp on CTHD.MaSanPham = sp.MaSanPham where MaHoaDon = '" + get + "'")
            cn.close()

            #Xóa hiện thị trên bảng chi tiết hóa đơn (nếu có)
            for item in self.tree4a.get_children():
                self.tree4a.delete(item)

            #Đưa dữ liệu leen bảng chi tiết hóa đơn
            a = 0
            for i in hd1:
                a += 1
                self.tree4a.insert(parent="", index='end', text=a, value=(i[0], i[1], i[2], i[3]))

        #Nếu không tìm thấy hóa đơn thì thông báo trên bảng hóa đơn
        else:
            #Xóa hiện thị trên bảng chi tiết hóa đơn (nếu có)
            for item in self.tree4a.get_children():
                self.tree4a.delete(item)
            #Thông báo không tìm thấy trên bảng hóa đơn
            self.tree4.insert(parent="", index='end', text=1, value=('Không tìm thấy',"","","",""))


    #4.3 Chức năng delete_hd, tab Hóa đơn: Xóa hóa đơn và chi tiết hóa đơn
    def delete_hd(self):
        #Lấy mã hóa đơn từ giao diện
        ma_hd = self.ma_hd_input.get()

        #Kiểm tra xem mã hóa đơn có tồn tại
        check = 0
        for row in UI.rows1:
            if row[0] == ma_hd:
                check = 1

        #Nếu mã hóa đơn tồn tại thì xóa
        if check:
            cn = MSSQLConnection()
            cn.connect()
            cn.delete("delete from CTHD where MaHoaDon = '" + ma_hd + "'")
            cn.delete("delete from HOADON where MaHoaDon = '" + ma_hd + "'")
            cn.close()

            #Tự động load CSDL lên bảng hóa đơn sau khi xóa hóa đơn
            self.load_hd()

            #Xóa hiển thị cũ trên bảng chi tiết hóa đơn
            for item in self.tree4a.get_children():
                self.tree4a.delete(item)

        #Nếu không tìm thấy mã hóa đơn thì chỉ cần thông báo trên bảng hóa đơn
        else:
            # Xóa hiện thị trên bảng chi tiết hóa đơn (nếu có)
            for item in self.tree4a.get_children():
                self.tree4a.delete(item)
            # Thông báo không tìm thấy trên bảng hóa đơn
            for item in self.tree4.get_children():
                self.tree4.delete(item)
            self.tree4.insert(parent="", index='end', text=1, value=("Không tìm thấy","","","",""))


    #5 Chức năng refresh_kh, tab Khách hàng: Đưa dữ liệu vào mảng
    def refresh_kh(self):
        cn = MSSQLConnection()
        cn.connect()
        UI.rows2 = cn.query("select * from KHACHHANG")
        cn.close()


    #5.1 Chức năng load_kh, tab Khách hàng: Đưa dữ liệu lên giao diện
    def load_kh(self):
        for item in self.tree5.get_children():
            self.tree5.delete(item)

        self.refresh_kh()

        for i in UI.rows2:
            self.tree5.insert(parent="", index='end', value=(i[0], i[1], i[2]))


    #5.2 Chức năng top5_kh, tab Khách hàng: Dùng câu truy vấn dưới CSDL để tìm 5 khách hàng mua hnagf nhiều nhất
    def top5_kh(self):
        for item in self.tree5.get_children():
            self.tree5.delete(item)

        cn = MSSQLConnection()
        cn.connect()
        top5 = cn.query("select top 5 hd.MaKhachHang, kh.HoTen, kh.SoDienThoai, sum(TriGia) as DoanhThu "
                        "from hoadon hd join khachhang kh on hd.MaKhachHang = kh.MaKhachHang "
                        "where hd.MaKhachHang is not null "
                        "group by hd.MaKhachHang, kh.HoTen, kh.SoDienThoai "
                        "order by DoanhThu desc")
        cn.close()

        for i in top5:
            self.tree5.insert(parent="", index='end', value=(i[0], i[1], i[2], i[3]))


    #5.3 Chức năng find_kh, tab Khách hàng: Tìm mã khách hàng
    def find_kh(self):
        self.refresh_kh()
        check = 0
        index = 0
        ma_kh = self.makh_kh_input.get()
        for row in UI.rows2:
            if row[0] == ma_kh:
                check = 1
                index = UI.rows2.index(row)

        for item in self.tree5.get_children():
            self.tree5.delete(item)

        if check:
            self.tree5.insert(parent="", index='end', text=1,
                             value=(UI.rows2[index][0], UI.rows2[index][1], UI.rows2[index][2]))
        else:
            self.tree5.insert(parent="", index='end', text=1, value=('Không tìm thấy', "", ""))



    #5.4 Chức năng delete_kh, tab Khách hàng: Xóa khách hàng
    def delete_kh(self):
        check = 0
        ma_kh = self.makh_kh_input.get()

        for row in UI.rows2:
            if row[0] == ma_kh:
                check = 1

        if check:
            cn = MSSQLConnection()
            cn.connect()
            cn.delete("delete from KHACHHANG where MaKhachHang = '" + ma_kh + "'")
            cn.close()
            self.load_kh()
        else:
            for item in self.tree1.get_children():
                self.tree5.delete(item)
            self.tree5.insert(parent="", index='end', value=("Không tồn tại", "", ""))


    #5.5 Chức năng add_kh, tab Khách hàng: Thêm khách hàng mới
    def add_kh(self):
        ten_kh = self.tenkh_kh_input.get()
        sdt_kh = self.sdtkh_kh_input.get()

        a = len(UI.rows2) - 1
        c = int((UI.rows2[a][0][2] + UI.rows2[a][0][3])) + 1
        b = 'KH' + str(c)

        cn = MSSQLConnection()
        cn.connect()
        cn.insert("insert into KHACHHANG values('" + b + "','" + ten_kh + "','" + sdt_kh + "')")
        cn.close()
        self.load_kh()


if __name__ == '__main__':
    ui = UI()
    ui.mainloop()