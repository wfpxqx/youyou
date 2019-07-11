CREATE DATABASE [CASCRS_VOUCHER]
GO
USE [CASCRS_VOUCHER]
GO
/****** Object:  Table [dbo].[CodeContrast]    Script Date: 2018-9-27 16:25:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CodeContrast](
	[autoId] [uniqueidentifier] NOT NULL,
	[middleCode] [nvarchar](50) NULL,
	[middleCodeName] [nvarchar](50) NULL,
	[targetCode] [nvarchar](50) NULL,
	[targetCodeName] [nvarchar](50) NULL,
	[Flag] [smallint] NULL,
 CONSTRAINT [PK_CodeContrast_1] PRIMARY KEY CLUSTERED 
(
	[autoId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[DeptItemContrast]    Script Date: 2018-9-27 16:25:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DeptItemContrast](
	[autoId] [uniqueidentifier] NOT NULL,
	[deptId] [nvarchar](50) NULL,
	[deptName] [nvarchar](50) NULL,
	[itemCClass] [nvarchar](50) NULL,
	[itemCName] [nvarchar](50) NULL,
	[itemId] [nvarchar](50) NULL,
	[itemName] [nvarchar](50) NULL,
 CONSTRAINT [PK_DeptItemContrast_1] PRIMARY KEY CLUSTERED 
(
	[autoId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[SystemUser]    Script Date: 2018-9-27 16:25:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[SystemUser](
	[userId] [nvarchar](50) NOT NULL,
	[userDisplayName] [nvarchar](50) NULL,
	[userPassword] [nvarchar](50) NULL,
 CONSTRAINT [PK_SystemUser] PRIMARY KEY CLUSTERED 
(
	[userId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
INSERT INTO dbo.SystemUser (userId, userDisplayName, userPassword) VALUES ('admin', '系统管理员', 'admin')
GO
