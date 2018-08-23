CREATE PROCEDURE [dbo].[Thiet_Bi_List]
AS
BEGIN
	SELECT [ID]
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
		  ,[Ma_Nhom]
		  ,[Ma_Chi_Tiet]
  FROM [EQUIP].[dbo].[Thiet_Bi] as Thiet_Bi
  inner join [EQUIP].[dbo].CS_tbPhong_Ban as Phong_Ban
  on Phong_Ban.ID = Thiet_Bi.Phong_Ban
END
GO
