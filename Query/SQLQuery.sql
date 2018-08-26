ALTER PROCEDURE [dbo].[Thiet_Bi_List]
AS
BEGIN
	SELECT Thiet_Bi.ID
		  ,[Ten_Thiet_Bi]
		  ,[Phong_Ban]
		  ,[Vi_Tri]
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
END
GO

ALTER PROCEDURE [dbo].[Thiet_Bi_List_By_Condition]
	@Phong_Ban as int,
	@Ma_Nhom as int
AS
BEGIN
	SELECT Thiet_Bi.ID
		  ,[Ten_Thiet_Bi]
		  ,[Phong_Ban]
		  ,[Vi_Tri]
		  ,[Hinh_Anh]
		  ,[Ma_Thiet_Bi]
		  ,[Ghi_Chu_1]
		  ,[Start_Date]
		  ,[End_Date]
		  ,[Ghi_Chu_2]
		  ,[Don_Gia]
		  ,Thiet_Bi.[Ma_Nhom]
		  ,[Ma_Chi_Tiet]
  FROM [EQUIP].[dbo].[Thiet_Bi] as Thiet_Bi
  inner join [EQUIP].[dbo].CS_tbPhong_Ban as Phong_Ban
  on Phong_Ban.ID = Thiet_Bi.Phong_Ban
  where Thiet_Bi.Phong_Ban = @Phong_Ban and Thiet_Bi.Ma_Nhom = @Ma_Nhom order by Thiet_Bi.ID
END
GO


DBCC CHECKIDENT ('[EQUIP].[dbo].[Thiet_Bi]', RESEED, 0);
GO
