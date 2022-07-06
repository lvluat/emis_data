$sql_instance = 'SERVER\SQLSERVER2017'; #assume we use windows authentication
$sql_database = 'EMIS.NET'
$sql_username = 'ssrs'
$sql_pass = '9981+DuLich'
$sql_queries = @()

$csv_path = 'E:\Softs\Prog\Git\bin\emis_data\'
$csv_files = @()
$csv_delim = '~'

$temp_file = $($csv_path + 'temp.csv')

$indent = "  "

# 
Set-ExecutionPolicy RemoteSigned

##################

$csv_files += 'DM_Lop'

$sql_queries += "
SELECT
	MaLopHoc
	,lh.MaQuyUoc AS MaLophocQuyUoc
	,lh.MaNganh
	,ng.MaQuyUoc AS MaNganhQuyUoc
	,TenLopHoc
	,lh.MaKhoaHoc
	,kh.MaQuyUoc AS MaKhoaHocQuyUoc
	,lh.MaBacDaoTao
	,MaChuyenNganh
	,MaLoaiDaoTao
	,MaKhoa
	,NgayNhapHoc
	,NgayRaTruong
	,OK
	,ThuTien
	,MaQuyChe
	,HocPhiTinChi
	,NhomLopHocPhan
	,lh.MaQuyUoc
	,NamHocBD
	,PhongChinh
	,GVCN
	,MaBuoiHoc
FROM DM_LopHoc lh
	JOIN DM_Nganh ng ON ng.MaNganh = lh.MaNganh
	JOIN DM_KhoaHoc kh ON kh.MaKhoaHoc = lh.MaKhoaHoc
"

			
$csv_files += 'DM_LopHoc'

$sql_queries += "
SELECT
	MaLopHoc
	,lh.MaQuyUoc AS MaLophocQuyUoc
	,lh.MaNganh
	,ng.MaQuyUoc AS MaNganhQuyUoc
	,TenLopHoc
	,lh.MaKhoaHoc
	,kh.MaQuyUoc AS MaKhoaHocQuyUoc
	,lh.MaBacDaoTao
	,MaChuyenNganh
	,MaLoaiDaoTao
	,MaKhoa
	,NgayNhapHoc
	,NgayRaTruong
	,OK
	,ThuTien
	,MaQuyChe
	,HocPhiTinChi
	,NhomLopHocPhan
	,lh.MaQuyUoc
	,NamHocBD
	,PhongChinh
	,GVCN
	,MaBuoiHoc
FROM DM_LopHoc lh
	JOIN DM_Nganh ng ON ng.MaNganh = lh.MaNganh
	JOIN DM_KhoaHoc kh ON kh.MaKhoaHoc = lh.MaKhoaHoc
"	

    
$csv_files += 'DM_KhoaHoc'

$sql_queries += "
SELECT
	kh.MaKhoaHoc
	,kh.MaQuyUoc AS MaKhoaHocQuyUoc
	,kh.CachViet
	,kh.NamBatDau
	,kh.MaBacDaoTao
	,kh.MaNhomKhoaHoc
	,bdt.MaQuyChe
	,bdt.SoNamHoc
	,bdt.ThoiGian AS ThoiGianDaoTao
FROM DM_KhoaHoc kh
	INNER JOIN DM_BacDaoTao bdt ON bdt.MaBacDaoTao = kh.MaBacDaoTao    
"	

    
$csv_files += 'EM_DT_ChuongTrinhDaoTao'

$sql_queries += "
SELECT            
	c.MaKhoaHoc
	,kh.MaQuyUoc AS MaKhoaHocQuyUoc
	,c.MaNganh
	,nh.MaQuyUoc AS MaNganhQuyUoc
	,c.MaBacDaoTao
	,c.MaKhoa
	,c.MaChuongTrinhDaoTao
	,c.MaMonHoc
	,mh.TenMonHoc
	,c.NamHoc
	,c.HocKy
	,c.GioLyThuyet AS GioLyThuyet
	,c.GioTuHoc AS GioKiemTra
	,c.GioThucHanh AS GioThucHanh
	,c.GioThucTap AS GioTuHoc
FROM EM_DT_ChuongTrinhDaoTao c
	INNER JOIN DM_Nganh nh ON nh.MaNganh = c.MaNganh
	INNER JOIN DM_KhoaHoc kh ON kh.MaKhoaHoc = c.MaKhoaHoc
	INNER JOIN DM_MonHoc mh ON mh.MaMonHoc = c.MaMonHoc
"	

    
$csv_files += 'EM_DT_SoDiem'

$sql_queries += "
SELECT DISTINCT
	sv.SinhVien
	,sv.MaLopHoc
	,sv.MaSinhVien
	,TRY_CAST((sv.NamSinh + '-' + sv.ThangSinh + '-' + sv.NgaySinh) AS date) AS NgaySinh
	,lhp.MaMonHoc
	,mh.TenMonHoc
	,ctdt.NamHoc
	,ctdt.HocKy
	,ctdt.KhoaLuanTN
	,ctdt.SoTinChi
	,lhp.MaChuongTrinhDaoTao
	,lh.MaQuyUoc AS MaLopQuyUoc
	,ng.MaQuyUoc AS MaNganhQuyUoc
	,lh.MaKhoaHoc	
	,kh.MaQuyUoc AS MaKhoaHocQuyUoc
	,sd.HS11
	,sd.HS12
	,sd.HS13
	,sd.HS14
	,sd.HS15
	,sd.HS16
	,sd.HS17
	,sd.HS18
	,sd.HS19
	,sd.HS110
	,sd.HS21
	,sd.HS22
	,sd.HS23
	,sd.HS24
	,sd.HS25
	,sd.HS26
	,sd.HS27
	,sd.HS28
	,sd.HS29
	,sd.HS210
	,sd.TBKiemTra
	,sd.Thi1
	,sd.Thi2
	,sd.Thi3
	,sd.TKM1
	,sd.TKM2
	,sd.TKM3
	,CASE
		WHEN Thi1 = - 1 THEN - 1
		WHEN ISNULL(sd.Thi1,0) >= ISNULL(sd.Thi2,0) AND ISNULL(sd.Thi1,0) >= ISNULL(sd.Thi3,0) THEN sd.Thi1
		WHEN ISNULL(sd.Thi2,0) >= ISNULL(sd.Thi3,0) THEN sd.Thi2
		ELSE sd.Thi3
	END AS DiemThiLN
	,CASE
		WHEN Thi1 = - 1 THEN - 1
		WHEN ISNULL(sd.TKM1,0) >= ISNULL(sd.TKM2,0) AND ISNULL(sd.TKM1,0) >= ISNULL(sd.TKM3,0) THEN sd.TKM1
		WHEN ISNULL(sd.TKM2,0) >= ISNULL(sd.TKM3,0) THEN sd.TKM2
		ELSE sd.TKM3
	END AS DiemLN
	,IIF(FORMAT(lh.NgayNhapHoc,'yy')=LEFT(sv.MaSinhVien,2),1,0) AS DungKhoaHoc
	,IIF(Thi1 IS NULL, 0, 1) + IIF(Thi2 IS NULL, 0, 1) + IIF(Thi3 IS NULL, 0, 1) AS SoLanThi
	,CAST(ISNULL(sd.Ngay1, sd.Ngay2) AS date) AS NgayNhapDiem
	,IIF(IIF(Thi1 IS NULL, 0, 1) + IIF(Thi2 IS NULL, 0, 1) + IIF(Thi3 IS NULL, 0, 1)=0, NULL,
		IIF(CASE
			WHEN Thi1 = - 1 THEN - 1
			WHEN ISNULL(sd.TKM1, 0) >= ISNULL(sd.TKM2, 0) AND ISNULL(sd.TKM1, 0) >= ISNULL(sd.TKM3, 0) THEN sd.TKM1
			WHEN ISNULL(sd.TKM2, 0) >= ISNULL(sd.TKM3, 0) THEN sd.TKM2
			ELSE sd.TKM3
		END>=5, 1, 0)
	) AS KetQuaMon
	,tn.DTBTN
	,tn.TBC
	,tn.TBCXH
	,tn.NamTN
	,tn.XepLoai_TN
	,tn.TotNghiep
	,tn.SoHieuBang
	,tn.SoVaoSo
	,(sv.Del | sv.DinhChi) AS DaNghi
	,sv.Del
	,sv.DinhChi
FROM EM_DT_SODIEM sd 
	INNER JOIN EM_HS_SinhVien sv ON sv.SinhVien = sd.SinhVien
	INNER JOIN EM_DT_LopHocPhan lhp ON lhp.MaLopHocPhan = sd.MaLopHocPhan
	INNER JOIN DM_LopHoc lh ON lh.MaLopHoc = sv.MaLopHoc
	INNER JOIN DM_Nganh ng ON ng.MaNganh = lh.MaNganh
	INNER JOIN EM_DT_ChuongTrinhDaoTao ctdt ON lhp.MaChuongTrinhDaoTao = ctdt.MaChuongTrinhDaoTao
	INNER JOIN DM_MonHoc mh ON ctdt.MaMonHoc = mh.MaMonHoc
	INNER JOIN DM_KhoaHoc kh ON kh.MaKhoaHoc = lh.MaKhoaHoc
	LEFT JOIN EM_DT_DSTotNghiep tn ON tn.SinhVien = sd.SinhVien
WHERE (lhp.LoaiNhom = 1 OR lhp.LoaiNhom = 2) AND lh.MaQuyUoc NOT LIKE N'%0-%'
ORDER BY CAST(ISNULL(sd.Ngay1, sd.Ngay2) AS date) ASC
"

    
	
$csv_files += 'EM_DT_ThoiKhoaBieu'

$sql_queries += "
SELECT DISTINCT
	t.MaGiaoVien
	,gv.HoDem
	,gv.Ten
	,gv.TenVietTat
	,gv.MaKhoa AS MaKhoaGV
	,lhp.MaLopHocPhan
	,lhp.MaChuongTrinhDaoTao
	,lhp.MaMonHoc
	,t.MaGiangDuong
	,t.Ngay
	,t.Tiet
	,t.NoiDungBaiGiang
	,t.NoiDungTKB
	,t.ThucHanh
	,t.ThucTap
	,CASE
		WHEN ctdtlhp.GioLyThuyet != 0 AND ctdtlhp.GioThucHanh = 0 THEN 'LT'
		WHEN ctdtlhp.GioLyThuyet = 0 AND ctdtlhp.GioThucHanh != 0 THEN 'TH'
		WHEN ctdtlhp.GioLyThuyet != 0 AND ctdtlhp.GioThucHanh != 0 THEN 'HH'
		ELSE 'ON'
	END AS LoaiTiet
	,ctdtlhp.NamHoc
	,ctdtlhp.HocKy
	,t.MaLopHocPhan
	,lhpn.MaLopHoc AS MaLopHocTKB
	,lh.MaKhoaHoc AS MaKhoaHocTKB
	,lh.MaKhoa AS MaKhoaTKB
	,lh.MaQuyUoc AS MaLopHocQuyUoc
	,ng.MaQuyUoc AS MaNganhQuyUoc
	,kh.MaQuyUoc AS MaKhoaHocQuyUoc
	,CASE WHEN Tiet <= 5 THEN 0 WHEN Tiet <= 10 THEN 1 ELSE 2 END AS Buoi
FROM dbo.EM_DT_ThoiKhoaBieu t
	LEFT JOIN EM_HS_GiaoVien gv ON gv.MaGiaoVien = t.MaGiaoVien
	INNER JOIN dbo.V_EM_CTDT_LOP_HP ctdtlhp ON t.MaLopHocPhan = ctdtlhp.MaLopHocPhan
	INNER JOIN dbo.EM_DT_LopHocPhan lhp ON t.MaLopHocPhan = lhp.MaLopHocPhan 
	INNER JOIN EM_DT_LopHocPhan_Nhom lhpn ON t.MaLopHocPhan = lhpn.MaLopHocPhan
	INNER JOIN DM_LopHoc lh ON lh.MaLopHoc = lhpn.MaLopHoc
	INNER JOIN DM_Nganh ng ON ng.MaNganh = lh.MaNganh
	INNER JOIN DM_KhoaHoc kh ON kh.MaKhoaHoc = lh.MaKhoaHoc
WHERE ctdtlhp.LoaiNhom=1 OR ctdtlhp.LoaiNhom=3
"


    
$csv_files += 'EM_HP_DuKienThu'

$sql_queries += "
SELECT
	sv.SinhVien
	,sv.MaSinhVien
	,sv.HoDem
	,sv.Ten
	,TRY_CAST(NamSinh + '-' + ThangSinh + '-' + NgaySinh AS date) AS NgaySinh
	,lh.MaKhoaHoc
	,kh.MaQuyUoc AS MaKhoaHocQuyUoc
	,kh.NamBatDau
	,lh.MaBacDaoTao
	,lh.MaNganh
	,ng.MaQuyUoc AS MaNganhQuyUoc
	,kt.MaNhomKhoanThu
	,dkt.MaKhoanThu
	,kt.TenKhoanThu
	,kt.BatBuoc
	,kt.MoTa
	,kt.Nam
	,kt.HocKy
	,kt.Nam-kh.NamBatDau AS TTT
	,PhaiNop
	,sv.DinhChi
	,sv.Del
FROM EM_HP_DuKienThu dkt
	INNER JOIN EM_HS_SinhVien sv ON sv.SinhVien = dkt.SinhVien
	INNER JOIN DM_KhoanThu kt ON kt.MaKhoanThu = dkt.MaKhoanThu
	INNER JOIN DM_LopHoc lh ON lh.MaLopHoc = sv.MaLopHoc
	INNER JOIN DM_Nganh ng ON ng.MaNganh = lh.MaNganh
	INNER JOIN DM_KhoaHoc kh ON kh.MaKhoaHoc = lh.MaKhoaHoc
"	


    
$csv_files += 'EM_HS_SinhVien'

$sql_queries += "
SELECT
	sv.SinhVien
	,sv.MaLopHoc
	,lh.MaKhoaHoc            
	,lh.MaKhoa
	,lh.MaBacDaoTao
	,lh.MaNganh
	,SoHoSo            
	,kh.MaQuyUoc AS MaKhoaHocQuyUoc
	,lh.MaQuyUoc AS MaLopHocQuyUoc
	,ng.MaQuyUoc AS MaNganhQuyUoc
	,MaSinhVien
	,HoDem
	,Ten
	,IIF(GioiTinh=1,N'Nam',N'Nữ') AS GioiTinh
	,TRY_CAST((NamSinh + '-' + RIGHT('00' + ThangSinh,2) + '-' + RIGHT('00' + NgaySinh, 2)) AS date) AS NgaySinh
	,Del
	,DinhChi
	,BaoLuu
	,TinhTrang
	,MaQuocTich
	,MaDanToc
	,MaTonGiao
	,QueQuan
	,NoiSinh
	,HoKhauThuongTru
	,DienThoai	
	,diachilienlac AS DiaChiLienLac
	,NamTotNghiepPTTH
	,TruongTHPT
	,TruongTotNghiepPhoThong
	,BangTotNghiepPhoThong AS BangTHPT
	,HBPT AS HocBaTHPT
	,BangTHCS
	,HocBaTHCS
	,HocBa
	,CASE MaKenhTuyenSinh
		WHEN '01' THEN '07'
		WHEN '02' THEN '08'
		WHEN '03' THEN '09'
		WHEN '04' THEN '10'
		ELSE MaKenhTuyenSinh
	END AS MaKenhTuyenSinh
	,CAST(IIF(SUBSTRING(sv.MaLopHoc,3,1)='0',0,1) AS bit) AS DaNhapHoc
FROM EM_HS_SinhVien sv
	JOIN DM_LopHoc lh ON lh.MaLopHoc = sv.MaLopHoc
	JOIN DM_Nganh ng ON ng.MaNganh = lh.MaNganh
	JOIN DM_KhoaHoc kh ON kh.MaKhoaHoc = lh.MaKhoaHoc
"	


    
$csv_files += 'EM_HS_TuyenSinh'

$sql_queries += "
SELECT
	ts.SinhVien_TS
	,MaKeHoachTuyenSinh
	,ts.MaKhoaHoc
	,ts.MaKenhTuyenSinh
	,kh.MaQuyUoc AS MaKhoaHocQuyUoc
	,ts.MaBacDaoTao
	,ts.MaNganh
	,ng.MaQuyUoc AS MaNganhQuyUoc
	,ts.SoHoSo
	,ts.HoDem
	,ts.Ten
	,IIF(ts.GioiTinh=1,N'Nam',N'Nữ') AS GioiTinh
	,TRY_CAST((ts.NamSinh + '-' + RIGHT('00' + ts.ThangSinh,2) + '-' + RIGHT('00' + ts.NgaySinh, 2)) AS date) AS NgaySinh
	,ts.DinhChi
	,ts.MaQuocTich
	,ts.MaDanToc
	,ts.MaTonGiao
	,ts.QueQuan
	,ts.NoiSinh
	,ts.HoKhauThuongTru
	,ts.DienThoai	
	,ts.DiaChiLienLac
	,ts.BangTNNghe
	,ts.NamTotNghiepPTTH
	,ts.TruongTotNghiepPhoThong
	,ts.MaTinhThanhPho
	,ts.MaHuyenTN
	,ts.GiayCNPT	
	,ts.BangTotNghiepPhoThong AS BangTHPT
	,ts.HBPT AS HocBaTHPT
	,ts.BangTHCS
	,ts.HocBaTHCS
	,NgayNopHS
	,DaNopHS
	,sv.MaLopHoc
	,IIF(sv.Del=1 OR ISNULL(sv.SoHoSo,0)=0 OR sv.MaLopHoc LIKE N'%0-%',0,1) AS DaNhapHoc
FROM EM_HS_TuyenSinh ts
	JOIN DM_Nganh ng ON ng.MaNganh = ts.MaNganh    
	JOIN DM_KhoaHoc kh ON kh.MaKhoaHoc = ts.MaKhoaHoc
	LEFT JOIN EM_HS_SinhVien sv ON sv.SoHoSo = ts.SoHoSo
"	


    
$csv_files += 'EM_HP_SoThu'

$sql_queries += "
SELECT 
	sv.MaLopHoc
	,lh.MaQuyUoc AS MaLopHocQuyUoc
	,lh.MaKhoaHoc
	,kh.MaQuyUoc AS MaKhoaHocQuyUoc
	,lh.MaNganh
	,ng.MaQuyUoc AS MaNganhQuyUoc
	,lh.MaBacDaoTao
	,kt.MaNhomKhoanThu
	,stct.MaKhoanThu
	,kt.Nam
	,kt.HocKy     
	,st.SinhVien       
	,sv.MaSinhVien            
	,sv.HoDem
	,sv.Ten
	,st.MaQuyenSo
	,st.SoPhieu
	,st.MaCoSo
	,st.NgayThu
	,st.LyDoThu
	,st.SoTien
	,ThuNganHang
	,st.DaXoa			
	,st.NguoiDung1
	,st.Ngay1
	,st.NguoiDung2
	,st.Ngay2
	,st.MAY1
	,st.MAY2            
FROM [EMIS.NET].[dbo].[EM_HP_SinhVienSoThu] st
	INNER JOIN EM_HP_SoThuChiTiet stct ON stct.SinhVien = st.SinhVien AND stct.MaQuyenSo = st.MaQuyenSo AND stct.SoPhieu = st.SoPhieu
	INNER JOIN DM_KhoanThu kt ON kt.MaKhoanThu = stct.MaKhoanThu
	INNER JOIN EM_HS_SinhVien sv ON sv.SinhVien = st.SinhVien
	INNER JOIN DM_LopHoc lh ON lh.MaLopHoc = sv.MaLopHoc
	INNER JOIN DM_KhoaHoc kh ON kh.MaKhoaHoc = lh.MaKhoaHoc
	INNER JOIN DM_Nganh ng ON ng.MaNganh = lh.MaNganh   
"	

    
$csv_files += 'EM_HP_KeHoachThu'

$sql_queries += "
SELECT
	lh.MaKhoaHoc
	,kh.MaQuyUoc AS MaKhoaHocQuyUoc
	,kh.NamBatDau
	,lh.MaBacDaoTao
	,lh.MaNganh
	,ng.MaQuyUoc AS MaNganhQuyUoc
	,kt.MaNhomKhoanThu
	,dkt.MaKhoanThu
	,kt.TenKhoanThu
	,kt.BatBuoc
	,kt.MoTa
	,kt.Nam
	,kt.HocKy
	,kt.Nam-kh.NamBatDau AS TTT
	,MAX(PhaiNop) As SoTien
FROM EM_HP_DuKienThu dkt
	INNER JOIN EM_HS_SinhVien sv ON sv.SinhVien = dkt.SinhVien
	INNER JOIN DM_KhoanThu kt ON kt.MaKhoanThu = dkt.MaKhoanThu
	INNER JOIN DM_LopHoc lh ON lh.MaLopHoc = sv.MaLopHoc
	INNER JOIN DM_Nganh ng ON ng.MaNganh = lh.MaNganh
	INNER JOIN DM_KhoaHoc kh ON kh.MaKhoaHoc = lh.MaKhoaHoc
GROUP BY
	lh.MaKhoaHoc
	,kh.MaQuyUoc
	,kh.NamBatDau
	,lh.MaBacDaoTao
	,lh.MaNganh
	,ng.MaQuyUoc
	,kt.MaNhomKhoanThu
	,dkt.MaKhoanThu
	,kt.TenKhoanThu			
	,kt.Nam
	,kt.HocKy    
	,kt.BatBuoc
	,kt.MoTa        
"	


$csv_files += 'EM_DT_TongHopDiem'

$sql_queries += "
SELECT
	MaLopHocQuyUoc
	,MaKhoaHocQuyUoc
	,MaNganhQuyUoc
	,MaMonHoc
	,MaChuongTrinhDaoTao
	,SUM(IIF(DungKhoaHoc=1 AND LanThi=0 AND DinhChi=0,1,0)) AS SoDungKhoaChuaThi
	,SUM(IIF(DungKhoaHoc=1 AND DinhChi=0, 1, 0)) AS SoDungKhoa
	,SUM(IIF(LanThi>0 AND DinhChi=0, 1, 0)) AS SoDaThi
	,SUM(IIF(DinhChi=0 AND KetQuaMon=1,1,0)) AS SoDat
	,SUM(IIF(DinhChi=0, 1, 0)) AS TongSo
	,MIN(Ngay) AS NgayDiem
	,MaLopHoc
	,MaKhoaHoc
	,MaNganh
FROM
	(
	SELECT
		sd.MaLopHocPhan
		,sv.SinhVien
		,sv.MaLopHoc
		,lh.MaKhoaHoc		
		,lh.MaNganh
		,lhp.MaChuongTrinhDaoTao		
		,lhp.MaMonHoc		
		,sv.MaSinhVien
		,lh.MaQuyUoc AS MaLopHocQuyUoc
		,ng.MaQuyUoc AS MaNganhQuyUoc
		,kh.MaQuyUoc AS MaKhoaHocQuyUoc
		,IIF(FORMAT(lh.NgayNhapHoc,'yy')=LEFT(sv.MaSinhVien,2),1,0) AS DungKhoaHoc
		,IIF(Thi1 IS NULL, 0, 1) + IIF(Thi2 IS NULL, 0, 1) + IIF(Thi3 IS NULL, 0, 1) AS LanThi
		,ISNULL(sd.Ngay1, sd.Ngay2) AS Ngay
		,sv.DinhChi
		,IIF(IIF(Thi1 IS NULL, 0, 1) + IIF(Thi2 IS NULL, 0, 1) + IIF(Thi3 IS NULL, 0, 1)=0, NULL,
			IIF(CASE
				WHEN Thi1 = - 1 THEN - 1
				WHEN ISNULL(sd.TKM1, 0) >= ISNULL(sd.TKM2, 0) AND ISNULL(sd.TKM1, 0) >= ISNULL(sd.TKM3, 0) THEN sd.TKM1
				WHEN ISNULL(sd.TKM2, 0) >= ISNULL(sd.TKM3, 0) THEN sd.TKM2
				ELSE sd.TKM3
			END>=5, 1, 0)
		) AS KetQuaMon
	FROM EM_DT_SODIEM sd 
		INNER JOIN EM_HS_SinhVien sv ON sv.SinhVien = sd.SinhVien
		INNER JOIN EM_DT_LopHocPhan lhp ON lhp.MaLopHocPhan = sd.MaLopHocPhan
		INNER JOIN DM_LopHoc lh ON lh.MaLopHoc = sv.MaLopHoc
		INNER JOIN DM_Nganh ng ON ng.MaNganh = lh.MaNganh
		INNER JOIN DM_KhoaHoc kh ON kh.MaKhoaHoc = lh.MaKhoaHoc
	WHERE lhp.LoaiNhom = 1 OR lhp.LoaiNhom = 2
	) d
GROUP BY 
	MaMonHoc
	,MaChuongTrinhDaoTao
	,MaLopHocPhan
	,MaLopHoc
	,MaKhoaHoc
	,MaNganh
	,MaKhoaHocQuyUoc
	,MaNganhQuyUoc
	,MaLopHocQuyUoc
"	

    
$csv_files += 'EM_DT_ThoiKhoaBieu_Buoi'

$sql_queries += "
SET DATEFIRST 7;
SELECT
	MaGiaoVien
	,HoDem
	,Ten
	,TenVietTat
	,MaKhoaGV
	,MaChuongTrinhDaoTao
	,MaMonHoc
	,TenMonHoc
	,MaGiangDuong
	,Ngay
	,DATEPART(dw, Ngay) AS Thu
	,CacTiet
	,NoiDungBaiGiang
	,NoiDungTKB
	,LoaiBuoi	
	,Buoi
	,CacLopHocCung
	,N'►' + IIF(Buoi=0,'S',IIF(Buoi=1,'C','T')) + ' (' + CacTiet + '): ' +
	MaGiangDuong + ' - ' + LoaiBuoi + CHAR(10) +
	N' •' + TenMonHoc + IIF(LEN(ISNULL(NoiDungBaiGiang, ''))=0, '', CHAR(10) + N' •' + NoiDungBaiGiang) + CHAR(10) +
	N' •' + LEFT(CacLopHocCung, LEN(CacLopHocCung)-1)
	AS TKB
	,Tuan
FROM
(SELECT
	MaGiaoVien
	,HoDem
	,Ten
	,TenVietTat
	,MaKhoaGV
	,MaChuongTrinhDaoTao
	,MaMonHoc
	,TenMonHoc
	,MaGiangDuong
	,Ngay
	,CAST(MIN(Tiet) as nvarchar) + N'→' + CAST(MAX(Tiet) AS nvarchar) AS CacTiet
	,NoiDungBaiGiang
	,NoiDungTKB
	,MIN(LoaiTiet) AS LoaiBuoi	
	,MIN(Buoi) AS Buoi
	,(SELECT DISTINCT
					lh.MaQuyUoc + ', ' AS [text()]
				FROM EM_DT_ThoiKhoaBieu t1
					INNER JOIN EM_DT_LopHocPhan lhp ON lhp.MaLopHocPhan = t1.MaLopHocPhan
					INNER JOIN EM_DT_LopHocPhan_Nhom lhpn ON lhpn.MaLopHocPhan = t1.MaLopHocPhan
					INNER JOIN DM_LopHoc lh ON lh.MaLopHoc = lhpn.MaLopHoc
				WHERE tkb.MaChuongTrinhDaoTao=lhp.MaChuongTrinhDaoTao
					AND tkb.Ngay=t1.Ngay
					AND tkb.MaGiaoVien=t1.MaGiaoVien
					AND tkb.MaGiangDuong=t1.MaGiangDuong
				FOR XML PATH ('')
	) AS CacLopHocCung
FROM (
		SELECT DISTINCT
			t.MaGiaoVien
			,gv.HoDem
			,gv.Ten
			,gv.TenVietTat
			,gv.MaKhoa AS MaKhoaGV
			,lhp.MaLopHocPhan
			,lhp.MaChuongTrinhDaoTao
			,lhp.MaMonHoc
			,TenMonHoc
			,t.MaGiangDuong
			,t.Ngay
			,t.Tiet
			,t.NoiDungBaiGiang
			,t.NoiDungTKB
			,t.ThucHanh
			,t.ThucTap
			,CASE
				WHEN ctdtlhp.GioLyThuyet != 0 AND ctdtlhp.GioThucHanh = 0 THEN 'LT'
				WHEN ctdtlhp.GioLyThuyet = 0 AND ctdtlhp.GioThucHanh != 0 THEN 'TH'
				WHEN ctdtlhp.GioLyThuyet != 0 AND ctdtlhp.GioThucHanh != 0 THEN 'HH'
				ELSE 'ON'
			END AS LoaiTiet
			,ctdtlhp.NamHoc
			,ctdtlhp.HocKy
			,lhpn.MaLopHoc AS MaLopHocTKB
			,lh.MaKhoaHoc AS MaKhoaHocTKB
			,lh.MaKhoa AS MaKhoaTKB
			,lh.MaQuyUoc AS MaLopHocQuyUoc
			,ng.MaQuyUoc AS MaNganhQuyUoc
			,kh.MaQuyUoc AS MaKhoaHocQuyUoc
			,CASE WHEN Tiet <= 5 THEN 0 WHEN Tiet <= 10 THEN 1 ELSE 2 END AS Buoi
		FROM dbo.EM_DT_ThoiKhoaBieu t
			LEFT JOIN EM_HS_GiaoVien gv ON gv.MaGiaoVien = t.MaGiaoVien
			INNER JOIN dbo.V_EM_CTDT_LOP_HP ctdtlhp ON t.MaLopHocPhan = ctdtlhp.MaLopHocPhan
			INNER JOIN dbo.EM_DT_LopHocPhan lhp ON t.MaLopHocPhan = lhp.MaLopHocPhan 
			INNER JOIN EM_DT_LopHocPhan_Nhom lhpn ON t.MaLopHocPhan = lhpn.MaLopHocPhan
			INNER JOIN DM_LopHoc lh ON lh.MaLopHoc = lhpn.MaLopHoc
			INNER JOIN DM_Nganh ng ON ng.MaNganh = lh.MaNganh
			INNER JOIN DM_KhoaHoc kh ON kh.MaKhoaHoc = lh.MaKhoaHoc
		WHERE ctdtlhp.LoaiNhom=1 OR ctdtlhp.LoaiNhom=3
	) tkb
GROUP BY 
	MaGiaoVien
	,HoDem
	,Ten
	,TenVietTat
	,MaKhoaGV
	,MaChuongTrinhDaoTao
	,MaMonHoc
	,TenMonHoc
	,MaGiangDuong
	,Ngay
	,Buoi
	,NoiDungBaiGiang
	,NoiDungTKB
	,NamHoc
	,HocKy
) a
	LEFT JOIN DM_TuanHoc tuan ON DATEADD(d, -((DATEPART(WEEKDAY, Ngay) - DATEPART(dw, '19000101') + 7) % 7), Ngay) = tuan.NgayBatDau
ORDER BY Ngay ASC, Buoi ASC
"

    
$csv_files += 'DM_TuanHoc'

$sql_queries += "
SELECT
	ID
	,NamHoc
	,HocKy
	,Tuan
	,NgayBatDau
	,MaNhomKhoaHoc
FROM DM_TuanHoc
"


# Title
cls
Write-Host "XUẤT DỮ LIỆU EMIS RA CSV"
Write-Host "LVL, 06/2022"

    
# Xuất dl ra file csv ma UTF8 65001
$log_file = $($csv_path + 'emis_data.log') # $($csv_path + (Get-Date).ToString('yyyMMddHHmmss') + '.log')
Write-Host ""
Write-Host $("Xuất " + $csv_files.Count + " files:")
$time = (Get-Date).ToString('HH:mm:ss dd/MM/yyyy')
Write-Host $("Bắt đầu: " + $time)
Write-Host ""

Out-File -FilePath $log_file -InputObject $("Export " + $csv_files.Count + " files:")
Out-File -FilePath $log_file -InputObject "" -Append
Out-File -FilePath $log_file -InputObject $("Start: " + $time) -Append
Out-File -FilePath $log_file -InputObject "" -Append

e:

cd $csv_path

for ($i = 0; $i -lt $csv_files.Count; $i++)
#for ($i = 0; $i -lt 2; $i++) # lệnh test với 2 query đầu
{
    Write-Host ""
    Write-Host $(($i+1).ToString('00') + "." + $file_name)

    Out-File -FilePath $log_file -InputObject "" -Append
    Out-File -FilePath $log_file -InputObject $(($i+1).ToString('00') + "." + $file_name) -Append

    # bcp $sql_queries[$i] queryout $file_name -S $sql_instance -U $sql_username -P $sql_pass -c -C 65001 -t $csv_delim | Select-String -Pattern 'rows copied', 'clock' >> $log_file
    sqlcmd -S $sql_instance -d $sql_database -U $sql_username -P $sql_pass -Q $sql_queries[$i] -p -u -W -w 65535 -s $csv_delim -o $temp_file
    $o = Get-Content -Path $temp_file     

    $file_name = $($csv_path + $csv_files[$i] + '.csv')

    # Lưu dữ liệu ra file csv    
    $o | Select-String -Pattern 'rows affected', 'Network packet size', 'xact', 'Clock Time', '--' -NotMatch > $temp_file
    $data = Get-Content -Path $temp_file
    $data | where {$_ -ne ""}  > $file_name
    #Out-File -FilePath $file_name -InputObject $data -Encoding utf8

    $f_name = Split-Path $file_name -Leaf
    ..\git add $f_name    

    # Log    
    
    $ts = ($o | Select-String -Pattern 'rows affected').ToString()
    $ts = $indent + $ts.ToString().Substring(1, $ts.ToString().Length-2) + "`n" + $indent + ($o | Select-String -Pattern 'Clock Time').ToString()
    #Write-Host $($ts.ToString().Substring(1, $ts.ToString().Length-2))

    #$ts = $o | Select-String -Pattern 'Clock Time'
    Out-File -FilePath $log_file -InputObject $ts -Append
    Write-Host $ts
}

$time = (Get-Date).ToString('HH:mm:ss dd/MM/yyyy')
Out-File -FilePath $log_file -InputObject "" -Append
Out-File -FilePath $log_file -InputObject $("Done: " + $time) -Append

$f_name = Split-Path $log_file -Leaf
..\git add $f_name

..\git commit -m "$time"

..\git push --force origin main

Write-Host ""
Write-Host $("===> Xong! " + $time)

Remove-Item  $temp_file