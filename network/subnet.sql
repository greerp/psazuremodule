USE [pghiscox]
GO

/****** Object:  Table [dbo].[subnet]    Script Date: 12/4/2015 6:58:17 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[subnet](
	[cidr] [nchar](19) NOT NULL,
	[range] [int] NOT NULL,
	[resourcegroup] [varchar](50) NULL,
	[allocated] [datetime] NULL,
 CONSTRAINT [PK_subnet] PRIMARY KEY CLUSTERED 
(
	[cidr] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)

GO

/*
   Simple stored procedure to get the next unallocated cidr that provides the minumum number of hosts
   Updates it to allocated and returns the CIDR to use to create the Azure 

*/


drop procedure pr_getresgroupcidr
go

create procedure pr_getresgroupcidr 
	@resgroup varchar(50),
	@range int
	
as begin
	declare @cidr varchar(50)

	if exists (select 1 from subnet where resourcegroup=@resgroup)
	begin
		throw 50000,'Resource Group Exists',1
		return 
	end


	select top 1 @cidr=cidr from subnet where resourcegroup is null and range>=@range order by range 

	if @cidr is not null begin
		update subnet set resourcegroup=@resgroup, allocated=getdate() where cidr=@cidr and resourcegroup is null
		select @cidr
		return
	end 
	else begin
		throw 50001,'No subnets available for the range specified',1
		select ''
		return 
	end
end

