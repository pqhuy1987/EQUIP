ALTER PROCEDURE [dbo].[Thiet_Bi_List]
AS
BEGIN
	SELECT Thiet_Bi.ID
		  ,[Ten_Thiet_Bi]
		  ,[Phong_Ban]
		  ,CS_tbViTri.CS_ViTri
		  ,[Hinh_Anh]
		  ,[Ma_Thiet_Bi]
		  ,[Ghi_Chu_1]
		  ,[Start_Date]
		  ,[End_Date]
		  ,[Ghi_Chu_2]
		  ,[Don_Gia]
		  ,Type
		  ,[Ma_Nhom]
		  ,[Ma_Chi_Tiet]
  FROM [EQUIP].[dbo].[Thiet_Bi] as Thiet_Bi
  inner join [EQUIP].[dbo].CS_tbPhong_Ban as Phong_Ban
  on Phong_Ban.ID = Thiet_Bi.Phong_Ban
  inner join [EQUIP].[dbo].CS_tbViTri as CS_tbViTri
  on CS_tbViTri.ID = Thiet_Bi.Vi_Tri
END
GO

ALTER PROCEDURE [dbo].[Thiet_Bi_List_By_Condition]
	@Phong_Ban as int,
	@Ma_Nhom as int
AS
BEGIN
	SELECT Thiet_Bi.ID
		  ,[Ten_Thiet_Bi]
		  ,Phong_Ban.Type
		  ,[Phong_Ban]
		  ,CS_tbViTri.CS_ViTri
		  ,[Hinh_Anh]
		  ,[Ma_Thiet_Bi]
		  ,[Ghi_Chu_1]
		  ,[Start_Date]
		  ,[End_Date]
		  ,[Ghi_Chu_2]
		  ,[Don_Gia]
		  ,Thiet_Bi.[Ma_Nhom]
		  ,Code_Group.Code
		  ,[Ma_Chi_Tiet]
  FROM [EQUIP].[dbo].[Thiet_Bi] as Thiet_Bi
  inner join [EQUIP].[dbo].CS_tbPhong_Ban as Phong_Ban
  on Phong_Ban.ID = Thiet_Bi.Phong_Ban
  inner join [EQUIP].[dbo].Code_Group as Code_Group
  on Code_Group.ID = Thiet_Bi.Ma_Nhom
  inner join [EQUIP].[dbo].CS_tbViTri as CS_tbViTri
  on CS_tbViTri.ID = Thiet_Bi.Vi_Tri
  where Thiet_Bi.Phong_Ban = @Phong_Ban and Thiet_Bi.Ma_Nhom = @Ma_Nhom order by Thiet_Bi.ID
END
GO

DBCC CHECKIDENT ('[EQUIP].[dbo].[Thiet_Bi]', RESEED, 0);
GO

--Điện thoại--
BEGIN
update Thiet_Bi
set
	Hinh_Anh = '5e9df1ff5d4dbc13e55c_2018_09_01_08_41_27.jpg'
  FROM [EQUIP].[dbo].[Thiet_Bi] as Thiet_Bi
  inner join [EQUIP].[dbo].CS_tbPhong_Ban as Phong_Ban
  on Phong_Ban.ID = Thiet_Bi.Phong_Ban
  inner join [EQUIP].[dbo].CS_tbViTri as CS_tbViTri
  on CS_tbViTri.ID = Thiet_Bi.Vi_Tri
  where Ma_Chi_Tiet = '51'

END

--Ghế NV--
BEGIN
update Thiet_Bi
set
  Hinh_Anh = '2beebee72149c0179958_2018_09_01_10_08_29.jpg'
  FROM [EQUIP].[dbo].[Thiet_Bi] as Thiet_Bi
  inner join [EQUIP].[dbo].CS_tbPhong_Ban as Phong_Ban
  on Phong_Ban.ID = Thiet_Bi.Phong_Ban
  inner join [EQUIP].[dbo].CS_tbViTri as CS_tbViTri
  on CS_tbViTri.ID = Thiet_Bi.Vi_Tri
  where Ma_Chi_Tiet = '13'

END

--Ban 0.6x12--
BEGIN
update Thiet_Bi
set
  Hinh_Anh = '465ba41934b7d5e98ca6_2018_09_01_10_19_53.jpg'
  FROM [EQUIP].[dbo].[Thiet_Bi] as Thiet_Bi
  inner join [EQUIP].[dbo].CS_tbPhong_Ban as Phong_Ban
  on Phong_Ban.ID = Thiet_Bi.Phong_Ban
  inner join [EQUIP].[dbo].CS_tbViTri as CS_tbViTri
  on CS_tbViTri.ID = Thiet_Bi.Vi_Tri
  where Ma_Chi_Tiet = '1'

END
